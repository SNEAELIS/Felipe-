import os
import time
import re
import sys
import asyncio
from datetime import timedelta
from typing import List, Dict, Optional, Any
from contextlib import asynccontextmanager

import pandas as pd
import aiofiles
import nest_asyncio

from playwright.async_api import TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError
from playwright.async_api import async_playwright, Page, Browser, Locator

from colorama import Fore, Style
from tqdm.asyncio import tqdm
from tqdm import tqdm as sync_tqdm

# Apply nest_asyncio to allow nested event loops (useful for Jupyter notebooks)
nest_asyncio.apply()


class BreakInnerLoop(Exception):
    pass


class AsyncPWRobo:
    def __init__(self, caminho_arquivo_saida: str, porta: int = 9222):
        self.porta = porta
        self.caminho_arquivo_saida = caminho_arquivo_saida
        self.playwright = None
        self.browser: Optional[Browser] = None
        self.context = None
        self.page: Optional[Page] = None

        # Initialize output dataframe and counter for batch saving
        self.unsaved_count = 0
        self.output_df = self._load_or_create_dataframe()

        # Semaphore for controlling concurrent access to browser
        self.browser_semaphore = asyncio.Semaphore(1)

        # Queue for saving operations
        self.save_queue = asyncio.Queue()
        self.save_task = None

    def _load_or_create_dataframe(self) -> pd.DataFrame:
        """Load existing dataframe or create new one"""
        if os.path.exists(self.caminho_arquivo_saida):
            try:
                return pd.read_excel(self.caminho_arquivo_saida)
            except Exception:
                return pd.DataFrame()
        return pd.DataFrame()

    async def initialize(self):
        """Initialize browser connection - EACH INSTANCE GETS ITS OWN"""
        self.playwright = await async_playwright().start()

        # Connect to Chrome instance on specific port
        cdp_url = f'http://localhost:{self.porta}'
        print(f"🔌 Port {self.porta}: Connecting to {cdp_url}")

        try:
            self.browser = await self.playwright.chromium.connect_over_cdp(cdp_url)

            if self.browser.contexts:
                self.context = self.browser.contexts[0]
                print(f"✅ Port {self.porta}: Using existing context")
            else:
                self.context = await self.browser.new_context()
                print(f"✅ Port {self.porta}: Created new context")

            self.page = None

            if self.context.pages:
                print(f"🔍 Port {self.porta}: Scanning {len(self.context.pages)} pages for valid one...")

                for i, page in enumerate(self.context.pages):
                    try:
                        url = page.url
                        print(f"  Page {i}: {url}")

                        # Skip chrome:// pages and newtab
                        if (url.startswith("chrome://") or
                                "newtab" in url or
                                url == "about:blank"):
                            print(f"  ⏭️  Skipping system page: {url}")
                            continue

                        # Check if it's our target domain
                        if "transferegov" in url:
                            self.page = page
                            print(f"  ✅ Found target page: {url}")
                            break
                        else:
                            # If not target but valid, keep as fallback
                            if not self.page:  # First valid non-chrome page
                                self.page = page
                                print(f"  ⚠️  Found fallback page (not target): {url}")

                    except Exception as e:
                        print(f"  ❌ Error checking page {i}: {e}")
                        continue

            # Enable resource blocking
            await self.block_rss()

            # Start save worker
            self.save_task = asyncio.create_task(self._save_worker())

            print(f"✅ Port {self.porta}: Fully initialized at {self.page.url}")

        except Exception as e:
            print(f"❌ Port {self.porta}: Failed to initialize: {e}")
            raise

    async def block_rss(self):
        """Block images, stylesheets, and fonts to make browsing faster"""

        async def route_handler(route, request):
            if request.resource_type in ["image", "stylesheet", "font"]:
                await route.abort()
            else:
                await route.continue_()

        await self.page.route('**/*', route_handler)
        print(f"✅ Resource blocking enabled for faster performance on port {self.porta}")

    async def consulta_proposta(self):
        """Navigate through the system tabs"""
        try:
            print(f"🔍 Port {self.porta}: Starting navigation")

            # First, make sure we're on the right page
            current_url = self.page.url
            print(f"📍 Port {self.porta}: Current URL: {current_url}")

            # Try to click the logo/home with longer timeout
            try:
                await self.page.locator('xpath=//*[@id="logo"]/a').click(timeout=5000)
                print(f"✅ Port {self.porta}: Clicked home/logo")
                await self.page.wait_for_load_state("networkidle")
            except Exception as e:
                print(f"⚠️ Port {self.porta}: Couldn't click logo: {e}")
                # Maybe we're already where we need to be

            # Wait a bit for any navigation to complete
            await asyncio.sleep(2)

            # Click menu items
            xpaths = [
                'xpath=//*[@id="menuPrincipal"]/div[1]/div[3]',
                'xpath=//*[@id="contentMenu"]/div[1]/ul/li[2]/a'
            ]

            for i, xpath in enumerate(xpaths):
                print(f"🖱️ Port {self.porta}: Clicking step {i + 1}")

                # Wait for element to be visible
                await self.page.locator(xpath).wait_for(state="visible", timeout=10000)
                await asyncio.sleep(0.5)  # Small pause

                # Click and wait for navigation
                await self.page.locator(xpath).click(timeout=10000)
                await self.page.wait_for_load_state("networkidle")

                print(f"✅ Port {self.porta}: Step {i + 1} completed")

            print(f"✅ Port {self.porta}: Navigation complete")

        except Exception as e:
            print(f"❌ Port {self.porta}: Navigation failed: {e}")
            # Take screenshot for debugging
            try:
                await self.page.screenshot(path=f"error_port_{self.porta}.png")
                print(f"📸 Port {self.porta}: Screenshot saved")
            except:
                pass
            raise

    async def campo_pesquisa(self, numero_processo: str):
        """Fill in the search field and access the desired program"""
        try:
            campo_pesquisa_locator = self.page.locator('xpath=//*[@id="consultarNumeroProposta"]')
            await campo_pesquisa_locator.fill(numero_processo)
            await campo_pesquisa_locator.press('Enter')

            try:
                acessa_item_locator = self.page.locator('xpath=//*[@id="tbodyrow"]/tr/td[1]/div/a')
                await acessa_item_locator.click(timeout=8000)
            except PlaywrightTimeoutError:
                print(f'❌ Process {numero_processo} not found on port {self.porta}')
                raise BreakInnerLoop
            except PlaywrightError as e:
                print(f'❌ Process {numero_processo} not found. Error: {type(e).__name__}')
                raise BreakInnerLoop

        except PlaywrightError as e:
            print(f'❌ Failed to search process {numero_processo}. Error: {type(e).__name__}')
            raise

    async def dados_proposta(self, tr: Locator) -> Optional[Dict[str, str]]:
        """Extract data from table row"""
        try:
            arvore_element = tr.locator('td.arvoreValores.arvoreExibe')

            # Extract tree values if present
            if await arvore_element.count() > 0:
                print(f"📊 Accessing value tree on port {self.porta}")
                all_text = await arvore_element.first.inner_text()
                lines = [line.strip() for line in all_text.split('\n') if line.strip()]

                result = {}
                for line in lines:
                    parts = line.split(maxsplit=2)
                    if len(parts) >= 3:
                        result[parts[2]] = ' '.join(parts[0:2])
                return result

            # Regular table data
            labels = await tr.locator('td.label').all_text_contents()
            fields = await tr.locator('td.field').all_text_contents()

            return {str(label).strip(): str(field).strip()
                    for label, field in zip(labels, fields) if label.strip()}

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error at line {exc_tb.tb_lineno} on port {self.porta}")
            print(f"❌ Unexpected error in dados_proposta: {type(e).__name__} - {str(e)[:100]}")

            # Navigate back on error
            await self.page.locator('xpath=//*[@id="lnkConsultaAnterior"]').click(timeout=2000)
            return None

    async def _save_worker(self):
        """Background worker for saving data to Excel"""
        while True:
            try:
                # Wait for save signal with timeout
                save_data = await asyncio.wait_for(self.save_queue.get(), timeout=1.0)
                if save_data == "STOP":
                    break

                if self.unsaved_count > 0:
                    await self._save_to_excel()

            except asyncio.TimeoutError:
                # Check if we need to save based on count
                if self.unsaved_count >= 10:
                    await self._save_to_excel()
                continue
            except Exception as e:
                print(f"❌ Error in save worker on port {self.porta}: {e}")

    async def _save_to_excel(self):
        """Actual Excel save operation"""
        try:
            # Run in thread pool to avoid blocking
            loop = asyncio.get_event_loop()
            await loop.run_in_executor(
                None,
                lambda: self.output_df.to_excel(self.caminho_arquivo_saida, index=False)
            )
            print(
                f"💾 Port {self.porta}: Saved {self.unsaved_count} records to {os.path.basename(self.caminho_arquivo_saida)}")
            self.unsaved_count = 0
        except Exception as e:
            print(f"❌ Error saving on port {self.porta}: {type(e).__name__}\n{str(e)[:100]}")

    async def save_data(self):
        """Queue data for saving"""
        await self.save_queue.put("SAVE")

    async def mark_as_done(self, raw_data: List[Optional[Dict]], numero_processo: str):
        """Safely mark row as done in the DataFrame"""

        def merge_data(raw_data: List[Optional[Dict]]) -> Dict:
            """Merge partial dictionaries into complete records"""
            final_record = {}

            for partial_dict in raw_data:
                if not partial_dict:
                    continue

                for key, value in partial_dict.items():
                    original_key = str(key).strip()
                    new_key = original_key
                    counter = 1

                    while new_key in final_record:
                        new_key = f"{original_key}_{counter}"
                        counter += 1

                    final_record[new_key] = value

            return final_record

        def sanitize_txt(txt: Any) -> str:
            """Remove invalid XML characters"""
            if isinstance(txt, str):
                return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', txt)
            return str(txt) if txt is not None else ""

        row_data = merge_data(raw_data)
        row_data = {k: sanitize_txt(v) for k, v in row_data.items()}

        try:
            df = self.output_df

            if not df.empty:
                id_col = df.columns[0]
                mask = df[id_col].astype(str).str.strip() == str(numero_processo).strip()

                if mask.any():
                    idx = df.index[mask][0]
                    for key, value in row_data.items():
                        if key not in df.columns:
                            df[key] = pd.NA
                        df.at[idx, key] = value
                    print(f"✅ Updated existing process: {numero_processo} on port {self.porta}")
                else:
                    df_update = pd.DataFrame([row_data])
                    self.output_df = pd.concat([df, df_update], ignore_index=True)
                    print(f"✅ Appended new process: {numero_processo} on port {self.porta}")
            else:
                self.output_df = pd.DataFrame([row_data])
                print(f"✅ Created new record: {numero_processo} on port {self.porta}")

            self.unsaved_count += 1

            if self.unsaved_count >= 10:
                await self.save_data()

        except Exception as err:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error at line {exc_tb.tb_lineno} on port {self.porta}")
            print(f"❌ Error in mark_as_done: {type(err).__name__}\n{str(err)[:100]}")
            return False

    async def loop_de_pesquisa(self, numero_processo: str):
        """Main research loop for a single process"""
        print(f'🔍 Starting extraction for {numero_processo} on port {self.porta}')

        xpath = '//*[@id="alterar"]/div/form/table'

        try:
            await self.campo_pesquisa(numero_processo=numero_processo)
            await self.page.locator(xpath).wait_for(state='visible')

            rows = await self.page.locator(xpath).locator('tr').all()
            raw_row_data = []

            for row in rows:
                row_data = await self.dados_proposta(row)
                if row_data:
                    raw_row_data.append(row_data)

            await self.mark_as_done(numero_processo=numero_processo, raw_data=raw_row_data)
            await self.page.locator('xpath=//*[@id="breadcrumbs"]/a[2]').click()

            print(f"{'✨' * 3}📦 Data collected for {numero_processo} on port {self.porta} {'✨' * 3}")

        except PlaywrightTimeoutError as t:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f'TIMEOUT on port {self.porta}: {str(t)[:50]}')
            print(f"Error at line: {exc_tb.tb_lineno}")
            await self.consulta_proposta()
            return False
        except Exception as erro:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error at line {exc_tb.tb_lineno} on port {self.porta}")
            print(f'❌ Failed: {type(erro).__name__}\n{str(erro)[:100]}...')
            await self.consulta_proposta()
            raise BreakInnerLoop

    async def cleanup(self):
        """Clean up resources"""
        if self.save_task:
            await self.save_queue.put("STOP")
            await self.save_task

        if self.browser:
            await self.browser.close()
        if self.playwright:
            await self.playwright.stop()

    @staticmethod
    def extrair_dados_excel(caminho_arquivo_fonte: str, filter_: bool = False) -> Optional[pd.DataFrame]:
        """Extract data from Excel file"""
        try:
            complete_data_frame = pd.read_excel(caminho_arquivo_fonte, dtype=str, sheet_name=None)
            sheet_names_list = list(complete_data_frame.keys())

            print("✅ Sheet Names Found:")
            for name in sheet_names_list:
                print(f"- {name}")

            data_frame = complete_data_frame['Sheet1']

            if filter_:
                data_frame = data_frame[data_frame['Nº Proposta'].astype(str).str.contains('/', na=False)]
                data_frame = data_frame.drop_duplicates()

            print(f"✅ Loaded {len(data_frame)} rows from Excel.")
            return data_frame

        except Exception as e:
            print(f"🤷‍♂️❌ Error reading Excel: {os.path.basename(caminho_arquivo_fonte)}.\n"
                  f"Error: {type(e).__name__}\n{str(e)}")
            return None

    @staticmethod
    def fix_prop_num(numero_proposta) -> Optional[str]:
        """Fix proposal number format"""
        if pd.isna(numero_proposta):
            return None

        numero_proposta = str(numero_proposta)

        if '_' in numero_proposta:
            if '/' in numero_proposta:
                numero_proposta = numero_proposta.replace('_', '/')
                left, right = numero_proposta.split('/', 1)
                fixed = f"{left.zfill(6)}/{right}"

                if re.match(r'^\d{6}/\d{4}', fixed):
                    return fixed
                return None
        else:
            if '/' in numero_proposta:
                left, right = numero_proposta.split('/', 1)
                fixed = f"{left.zfill(6)}/{right}"

                if re.match(r'^\d{6}/\d{4}', fixed):
                    return fixed
                return None
        return None


class AsyncScraperManager:
    """Manages multiple async scraper instances"""

    def __init__(self, portas: List[int], arquivo_fonte: str, arquivo_saida_base: str):
        self.portas = portas
        self.arquivo_fonte = arquivo_fonte
        self.arquivo_saida_base = arquivo_saida_base
        self.robos: List[AsyncPWRobo] = []

    async def initialize_all(self):
        """Initialize all scraper instances with staggered starts"""
        for i, porta in enumerate(self.portas):
            saida = self.arquivo_saida_base.replace('.xlsx', f'_{porta}.xlsx')
            robo = AsyncPWRobo(saida, porta)

            # Stagger initialization by 2 seconds per instance
            await asyncio.sleep(i * 2)  # 0s, 2s, 4s, 6s...

            await robo.initialize()
            self.robos.append(robo)
            print(f"✅ Initialized port {porta} (instance {i + 1}/{len(self.portas)})")

    async def distribute_work(self, items: List, chunk_size: Optional[int] = None):
        """Distribute work items across scrapers"""
        if chunk_size is None:
            chunk_size = len(items) // len(self.robos) + 1

        chunks = [items[i:i + chunk_size] for i in range(0, len(items), chunk_size)]

        tasks = []
        for robo, chunk in zip(self.robos, chunks):
            if chunk:
                task = asyncio.create_task(self._process_chunk(robo, chunk))
                tasks.append(task)

        return await asyncio.gather(*tasks, return_exceptions=True)

    async def _process_chunk(self, robo: AsyncPWRobo, chunk: List):
        """Process a chunk of items with a specific scraper"""
        try:
            await robo.consulta_proposta()

            for item in tqdm(chunk, desc=f"Port {robo.porta}", position=robo.porta % 10):
                numero_processo = robo.fix_prop_num(item)

                if not numero_processo:
                    continue

                try:
                    await robo.loop_de_pesquisa(numero_processo=numero_processo)
                except BreakInnerLoop:
                    await robo.save_data()
                    break

            await robo.save_data()

        except Exception as e:
            print(f"❌ Error in port {robo.porta}: {e}")
        finally:
            await robo.cleanup()

    async def cleanup_all(self):
        """Clean up all scraper instances"""
        for robo in self.robos:
            await robo.cleanup()


def get_number_part(proposal: str) -> str:
    """Extract the number before the slash and remove leading zeros"""
    if '/' in proposal:
        return proposal.split('/')[0].lstrip('0') or '0'
    return proposal.lstrip('0') or '0'


async def main_async():
    """Main async function"""

    # Configuration
    arquivo_fonte = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Propostas_Extraidas_filtradas.xlsx"

    arquivo_saida_base = (r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência "
                          r"Social\Teste001\resultado_aba_dados.xlsx")

    # Ports to use (make sure Chrome is running with these debugging ports)
    portas = [9222, 9224, 9226, 9228]  # Add your ports here

    reset = input('Do you want to reset all files?')
    if reset.lower() == 'y':
        for porta in portas:
            path = arquivo_saida_base.replace('.xlsx', f'_{porta}.xlsx')

            if not os.path.exists(path):
                print(f"⚠️ Path not found: {path}")
                continue

            elif os.path.isfile(path):
                os.remove(path)
                print(f"🗑️ File deleted: {path}")
            else:
                print(f"⚠️ Unknown type (not file/dir): {path}")

    # Load the full DataFrame
    df = AsyncPWRobo.extrair_dados_excel(arquivo_fonte, filter_=True)
    if df is None:
        print("❌ Failed to load Excel file")
        return
    print(f"\n📊 Total proposals to process: {len(df)}")

    # SPLIT THE DATAFRAME INTO 4 EQUAL PARTS
    chunk_size = len(df) // len(portas)
    chunks = []

    for i, porta in enumerate(portas):
        output_file = arquivo_saida_base.replace('.xlsx', f'_{porta}.xlsx')

        # Start with the full dataframe for this port's chunk calculation
        chunk_size = len(df) // len(portas)
        start_idx = i * chunk_size
        end_idx = start_idx + chunk_size if i < len(portas) - 1 else len(df)

        # Get the raw chunk for this port
        raw_chunk = df.iloc[start_idx:end_idx].copy()
        print(f"\n📦 Port {porta}: Raw chunk has {len(raw_chunk)} proposals (rows {start_idx}-{end_idx})")

        # Filter out already processed items from this specific port's file
        if os.path.exists(output_file):
            print(f"🔍 Port {porta}: Checking existing file: {os.path.basename(output_file)}")

            # Load the existing output file
            df_existing = pd.read_excel(output_file)

            # Find completed programs (with Valor Global filled)
            if 'Valor Global' in df_existing.columns and 'Número da Proposta' in df_existing.columns:
                finished_programs = df_existing[
                    df_existing['Valor Global'].notna() &
                    (df_existing['Valor Global'] != '')
                    ]

                # Get list of completed proposal numbers
                to_jump = finished_programs['Número da Proposta'].tolist()

                if to_jump:
                    print(f"✅ Port {porta}: Found {len(to_jump)} already completed proposals")

                    # Extract number parts for comparison
                    finished_numbers = {get_number_part(p) for p in to_jump if len(p.split('/')) <= 6}
                    source_numbers = raw_chunk.iloc[:, 0].apply(get_number_part)

                    # Create filter mask
                    filter_mask = source_numbers.isin(finished_numbers)
                    filtered_out = filter_mask.sum()

                    # Apply filter
                    raw_chunk = raw_chunk[~filter_mask]

                    print(f"🎯 Port {porta}: Filtered out {filtered_out} already completed proposals")
                else:
                    print(f"📭 Port {porta}: No completed proposals found in output file")
            else:
                print(f"⚠️ Port {porta}: Output file missing required columns")
        else:
            print(f"🆕 Port {porta}: No existing output file, processing all {len(raw_chunk)} proposals")

        chunks.append(raw_chunk)
        print(f"📊 Port {porta}: Final chunk has {len(raw_chunk)} proposals to process")

    # Create manager with pre-split chunks
    manager = AsyncScraperManager(portas, arquivo_fonte, arquivo_saida_base)

    try:
        # Initialize all scrapers
        await manager.initialize_all()

        # Process each chunk on its assigned port
        tasks = []
        for robo, chunk in zip(manager.robos, chunks):
            if not chunk.empty:
                items = chunk.iloc[:, 0].tolist()  # Get first column values
                print(f"\n🎯 Port {robo.porta} assigned {len(items)} proposals")
                task = asyncio.create_task(manager._process_chunk(robo, items))
                tasks.append(task)

        # Wait for all to complete
        start_time = time.time()
        results = await asyncio.gather(*tasks, return_exceptions=True)

        # Report results
        elapsed = time.time() - start_time
        successful = sum(1 for r in results if not isinstance(r, Exception))
        failed = sum(1 for r in results if isinstance(r, Exception))

        print(f"\n{'=' * 60}")
        print(f"🎉 ALL INSTANCES COMPLETED!")
        print(f"⏱️  Total time: {timedelta(seconds=int(elapsed))}")
        print(f"✅ Successful: {successful}/{len(portas)} instances")
        print(f"❌ Failed: {failed}/{len(portas)} instances")

    except KeyboardInterrupt:
        print("\n🛑 Script stopped by user")
    except Exception as e:
        print(f"❌ Fatal error: {e}")
    finally:
        # Cleanup
        await manager.cleanup_all()


def main():
    """Entry point - runs the async main"""
    asyncio.run(main_async())


if __name__ == "__main__":
    main()