import os
import time
import re
import sys
import math
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

os.system('cls' if os.name == 'nt' else 'clear')

# Apply nest_asyncio to allow nested event loops
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

            if not self.page:
                raise RuntimeError(f"Port {self.porta}: No usable browser page found in the connected context.")

            # Start save worker
            self.save_task = asyncio.create_task(self._save_worker())

            print(f"✅ Port {self.porta}: Fully initialized")

        except Exception as e:
            print(f"❌ Port {self.porta}: Failed to initialize: {e}")
            raise

    async def safe_click(self, xpath: str, timeout: float = 10):
        """Safely click an element"""
        try:
            locator = self.page.locator(f'xpath={xpath}')
            await locator.wait_for(state="visible", timeout=timeout * 1000)
            await locator.click(timeout=timeout * 1000)
            await self.page.wait_for_load_state("networkidle")
            return True
        except Exception as e:
            print(f"⚠️ Port {self.porta}: Click failed on {xpath}: {e}")
            return False

    async def consulta_proposta(self):
        """Navigate through the system tabs to initial page"""
        try:
            print(f"🔍 Port {self.porta}: Starting navigation to initial page")
            
            # Try to click logo to reset to initial page
            try:
                logo = self.page.locator('xpath=//*[@id="logo"]/a')
                if await logo.count() > 0:
                    await logo.click(timeout=5000)
                    await self.page.wait_for_load_state("networkidle")
                    print(f"✅ Port {self.porta}: Clicked logo to reset")
            except Exception as e:
                print(f"⚠️ Port {self.porta}: Couldn't click logo (might already be on initial page): {e}")

            # Wait a bit for any navigation to complete
            await asyncio.sleep(2)

            # Navigate through menu
            xpaths = [
                '//*[@id="menuPrincipal"]/div[1]/div[3]',
                '//*[@id="contentMenu"]/div[1]/ul/li[2]/a'
            ]

            for i, xpath in enumerate(xpaths):
                print(f"🖱️ Port {self.porta}: Clicking menu step {i + 1}")
                
                if not await self.safe_click(xpath):
                    print(f"❌ Port {self.porta}: Failed to click menu item {i + 1}")
                    # Take screenshot for debugging
                    try:
                        await self.page.screenshot(path=f"error_menu_port_{self.porta}.png")
                        print(f"📸 Port {self.porta}: Screenshot saved")
                    except:
                        pass
                    raise Exception(f"Menu navigation failed at step {i + 1}")

                await asyncio.sleep(0.5)

            print(f"✅ Port {self.porta}: Navigation complete")

        except Exception as e:
            print(f"❌ Port {self.porta}: Navigation failed: {e}")
            raise

    async def acomp_fisc_tab(self):
        """Navigate to Acomp. e Fiscalização and Esclarecimentos tabs"""
        try:
            # Click on Acomp. e Fiscalização
            acomp_locator = self.page.locator("xpath=//div[contains(@class, 'button menu') and contains(text(), 'Acomp. e Fiscalização')]")
            await acomp_locator.wait_for(state="visible", timeout=10000)
            await acomp_locator.click()
            await self.page.wait_for_load_state("networkidle")
            
            # Click on Esclarecimentos
            esclarecimentos_locator = self.page.locator("#contentMenuInterno > div > div:nth-child(2) > ul > li:nth-child(3) > a")
            await esclarecimentos_locator.wait_for(state="visible", timeout=10000)
            await esclarecimentos_locator.click()
            await self.page.wait_for_load_state("networkidle")
            
            print(f'{"<" * 6} SUCESSO — Aba "Esclarecimento" acessada {">" * 6}'.center(80))
            
        except Exception as e:
            print(f"{Fore.RED}❌ Port {self.porta}: Error in acomp_fisc_tab: {e}{Style.RESET_ALL}")
            raise

    async def campo_pesquisa(self, numero_processo: str):
        """Fill in the search field and access the desired program"""
        try:
            if not self.page:
                raise RuntimeError("No page available for search")

            campo_pesquisa_locator = self.page.locator('xpath=//*[@id="consultarNumeroProposta"]')
            await campo_pesquisa_locator.wait_for(state="visible", timeout=10000)
            await campo_pesquisa_locator.fill(numero_processo)
            await campo_pesquisa_locator.press('Enter')
            
            # Wait for results to load
            await self.page.wait_for_load_state("networkidle")
            
            try:
                acessa_item_locator = self.page.locator('xpath=//*[@id="tbodyrow"]/tr/td[1]/div/a')
                await acessa_item_locator.click(timeout=8000)
                await self.page.wait_for_load_state("networkidle")
            except PlaywrightTimeoutError:
                print(f' Process number: {numero_processo}, not found on port {self.porta}.')
                raise BreakInnerLoop
            except PlaywrightError as e:
                print(f' Process number: {numero_processo}, not found. Error: {type(e).__name__}')
                raise BreakInnerLoop

        except PlaywrightError as e:
            print(f'❌ Port {self.porta}: Failed to search process {numero_processo}. Error: {type(e).__name__}')
            raise
        except Exception as e:
            print(f'❌ Port {self.porta}: campo_pesquisa crashed for {numero_processo}: {type(e).__name__}: {e}')
            raise

    async def dados_detalhamento(self) -> Dict[str, str]:
        """Extract data from detalhamento page"""
        tentativas = 0
        max_tentativas = 3
        
        while tentativas < max_tentativas:
            try:
                # Wait for labels to appear
                await self.page.locator('td.label').first.wait_for(state="visible", timeout=5000)
                
                # Extract labels and fields
                labels = await self.page.locator('td.label').all_text_contents()
                fields = await self.page.locator('td.field').all_text_contents()
                
                data = {str(l).strip(): str(f).strip() for l, f in zip(labels, fields) if l and l.strip()}
                
                # Check if empty (retry if page didn't load content)
                if not data:
                    print(f"{Fore.YELLOW}⚠️ Port {self.porta}: Page loaded but dictionary is empty. Retrying {tentativas+1}/{max_tentativas}...{Style.RESET_ALL}")
                    tentativas += 1
                    await asyncio.sleep(1.5)
                    continue
                
                # Click Voltar
                await self.page.locator('input[value="Voltar"]').click()
                await self.page.wait_for_load_state("networkidle")
                
                return data
                
            except (PlaywrightTimeoutError, PlaywrightError) as e:
                tentativas += 1
                print(f"{Fore.YELLOW}🔄 Port {self.porta}: {type(e).__name__} on attempt {tentativas}. Retrying...{Style.RESET_ALL}")
                await asyncio.sleep(0.5)
                
            except Exception as e:
                exc_type, exc_value, exc_tb = sys.exc_info()
                print(f"{Fore.RED}❌ Port {self.porta}: Hard Error on Line {exc_tb.tb_lineno}: {type(e).__name__}{Style.RESET_ALL}")
                break
                
        return {}  # Only returns empty after 3 failed attempts

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
                if self.unsaved_count >= 1:
                    await self._save_to_excel()
                continue
            except Exception as e:
                print(f"❌ Port {self.porta}: Error in save worker: {e}")

    async def _save_to_excel(self):
        """Actual Excel save operation"""
        try:
            # Run in thread pool to avoid blocking
            loop = asyncio.get_event_loop()
            await loop.run_in_executor(
                None,
                lambda: self.output_df.to_excel(self.caminho_arquivo_saida, index=False)
            )
            print(f"💾 Port {self.porta}: Saved {self.unsaved_count} records")
            self.unsaved_count = 0
        except Exception as e:
            print(f"❌ Port {self.porta}: Error saving: {type(e).__name__}\n{str(e)[:100]}")

    async def save_data(self):
        """Queue data for saving"""
        await self.save_queue.put("SAVE")

    async def search_loop_back(self):
        """Return to search page after processing"""
        try:
            await self.safe_click('//*[@id="menuInterno"]/div/div[4]')
            await self.safe_click('//*[@id="contentMenuInterno"]/div/div[1]/ul/li[6]/a')
            print(f"✅ Port {self.porta}: Returned to search page")
        except Exception as e:
            print(f"❌ Port {self.porta}: Error in search_loop_back: {e}")

    async def mark_as_done(self, raw_data_list: List[Dict], numero_processo: str):
        """Safely mark row as done in the DataFrame"""
        
        def sanitize_txt(txt: Any) -> str:
            """Remove invalid characters"""
            if isinstance(txt, str):
                return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', txt)
            return str(txt) if txt is not None else ""

        try:
            # Prepare the new rows
            new_rows_list = []
            for data_item in raw_data_list:
                if data_item:
                    sanitized_row = {str(k): sanitize_txt(v) for k, v in data_item.items()}
                    new_rows_list.append(sanitized_row)

            if not new_rows_list:
                return  # Nothing to add

            # Create a DataFrame from the new data
            new_data_df = pd.DataFrame(new_rows_list).astype(object)

            # Append to the main DataFrame
            if self.output_df.empty:
                self.output_df = new_data_df
            else:
                self.output_df = self.output_df.astype(object)
                self.output_df = pd.concat([self.output_df, new_data_df], ignore_index=True)

            print(f"✅ Port {self.porta}: Added {len(new_rows_list)} rows for process: {numero_processo}")

            # Batch saving logic
            self.unsaved_count += len(new_rows_list)
            if self.unsaved_count >= 1:
                await self.save_data()

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"{Fore.RED}❌ Port {self.porta}: DataFrame append failed on Line {exc_tb.tb_lineno}: {type(e).__name__}{Style.RESET_ALL}")

    async def click_page(self, page_number: int) -> bool:
        """Click on specific page number"""
        selectors = [f"//a[@title='Vá para a pág {page_number}']", f"//a[text()='{page_number}']"]
        for sel in selectors:
            try:
                await self.page.locator(f'xpath={sel}').click(timeout=2000)
                await self.page.wait_for_load_state("networkidle")
                return True
            except:
                continue
        return False

    async def loop_de_pesquisa(self, numero_processo: str):
        """Main research loop for a single process"""
        print(f'🔍 Port {self.porta}: Processing: {numero_processo}'.center(50, '-'))
        
        try:
            await self.campo_pesquisa(numero_processo)
            await self.acomp_fisc_tab()
        except:
            pass
        try:
            try:
                # Wait for esclarecimentos table to load
                await self.page.locator('#esclarecimentos').wait_for(state="visible", timeout=5000)
            except PlaywrightTimeoutError:
                print(f"{Fore.YELLOW}⚠️ Port {self.porta}: Esclarecimentos table not found for {numero_processo}. Marking as done with basic info.{Style.RESET_ALL}")
                await self.mark_as_done([{'Número da Proposta': numero_processo}], numero_processo)
                await self.consulta_proposta()
                return

            raw_row_data = []
            process_data = {'Número da Proposta': numero_processo}
            
            # Get row count
            rows = await self.page.locator('#esclarecimentos tbody tr').all()
            rows_count = len(rows)
            
            if rows_count > 0:
                # Get pagination info
                pg_info_text = await self.page.locator('//*[@id="esclarecimentos"]/span[1]').text_content()
                if pg_info_text:
                    match = re.search(r'(\d+)\s*\(', pg_info_text)
                    paginas = int(match.group(1)) if match else 1
                else:
                    paginas = 1
                
                for p in range(1, paginas + 1):
                    if p > 1:
                        await self.click_page(p)
                    
                    await self.page.locator('#esclarecimentos tbody tr').first.wait_for(state="visible", timeout=5000)
                    
                    # Process each row
                    current_rows = await self.page.locator('#esclarecimentos tbody tr').all()
                    for i in range(len(current_rows)):
                        if i > 0 and p > 1:
                            await self.click_page(p)
                            current_rows = await self.page.locator('#esclarecimentos tbody tr').all()
                        
                        # Click Detalhar button
                        detalhar_btn = current_rows[i].locator('a:has-text("Detalhar")')
                        await detalhar_btn.click()
                        
                        # Extract data
                        row_data = await self.dados_detalhamento()
                        if row_data:
                            raw_row_data.append({**process_data, **row_data})
                
                await self.mark_as_done(raw_row_data, numero_processo)
                await self.search_loop_back()
                print(f"✅ Port {self.porta}: Data collected for {numero_processo}")
            else:
                # No esclarecimentos found, just mark as done with basic info
                print('Callig mark_as_done with empty data since no esclarecimentos found')
                await self.mark_as_done([process_data], numero_processo)
                await self.search_loop_back()
                
        except BreakInnerLoop:
            raise
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"{Fore.RED}❌ Port {self.porta}: Loop crash on Line {exc_tb.tb_lineno}: {type(e).__name__}{Style.RESET_ALL}")
            print(f'Err msg: {str(e)[:100]}')

            # No esclarecimentos found, just mark as done with basic info
            print('Callig mark_as_done with empty data since no esclarecimentos found')
            await self.mark_as_done([process_data], numero_processo)
            await self.search_loop_back()
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
    def get_number_part(proposal: str) -> str:
        """Extract the number before the slash and remove leading zeros"""
        if '/' in proposal:
            return proposal.split('/')[0].lstrip('0') or '0'
        return proposal.lstrip('0') or '0'

    @staticmethod
    def extrair_dados_excel(caminho: str, filter_: bool = False) -> Optional[pd.DataFrame]:
        """Extract data from Excel file"""
        try:
            df = pd.read_excel(caminho, dtype=str)
            if filter_:
                df = df[df['Nº Proposta'].astype(str).str.contains('/', na=False)].drop_duplicates()
            print(f"✅ Loaded {len(df)} rows from Excel.")
            return df
        except Exception as e:
            print(f"🤷‍♂️❌ Error reading Excel: {os.path.basename(caminho)}.\nError: {type(e).__name__}\n{str(e)}")
            return None

    @staticmethod
    def fix_prop_num(numero_proposta) -> Optional[str]:
        """Fix proposal number format"""
        if pd.isna(numero_proposta):
            return None
        
        numero_proposta = str(numero_proposta).strip()
        
        # Remove any leading/trailing whitespace and handle underscores
        if '_' in numero_proposta:
            numero_proposta = numero_proposta.replace('_', '/')
        
        pattern = r'^\d{6}/\d{4}'
        
        if re.findall(pattern, numero_proposta):
            return numero_proposta
        else:
            if '/' in numero_proposta:
                parts = numero_proposta.split('/')
                if len(parts) == 2:
                    first_digits = re.sub(r'\D', '', parts[0])
                    second_digits = re.sub(r'\D', '', parts[1])
                    
                    if first_digits and second_digits:
                        first_padded = f"{int(first_digits):06d}"
                        return f"{first_padded}/{second_digits}"
        
        return None


class AsyncScraperManager:
    """Manages multiple async scraper instances"""
    
    def __init__(self, portas: List[int], arquivo_saida_base: str):
        self.portas = portas
        self.arquivo_saida_base = arquivo_saida_base
        self.robos: List[AsyncPWRobo] = []
    
    async def initialize_all(self):
        """Initialize all scraper instances with staggered starts"""
        for i, porta in enumerate(self.portas):
            saida = self.arquivo_saida_base.replace('.xlsx', f'_{porta}.xlsx')
            robo = AsyncPWRobo(saida, porta)
            
            # Stagger initialization by 2 seconds per instance
            if i > 0:
                await asyncio.sleep(2)
            
            await robo.initialize()
            self.robos.append(robo)
            print(f"✅ Initialized port {porta} (instance {i + 1}/{len(self.portas)})")
    
    async def _process_chunk(self, robo: AsyncPWRobo, items: List[str]):
        """Process a chunk of items with a specific scraper"""
        try:
            await robo.consulta_proposta()
            
            for numero_processo in tqdm(items, desc=f"Port {robo.porta}", position=robo.porta % 10):
                if not numero_processo:
                    continue
                
                try:
                    await robo.loop_de_pesquisa(numero_processo=numero_processo)
                except BreakInnerLoop:
                    await robo.save_data()
                    continue  # Try next proposal if one fails
            
            await robo.save_data()
            
        except Exception as e:
            print(f"❌ Port {robo.porta}: Error processing chunk: {e}")
    
    async def cleanup_all(self):
        """Clean up all scraper instances"""
        for robo in self.robos:
            await robo.cleanup()


async def main_async(to_jump: Optional[List[str]] = None):
    """Main async function"""
    
    # Configuration
    arquivo_fonte = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\Base de dados propostas\Propostas_Extraidas_filtradas.xlsx"
    arquivo_saida_base = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\resultado_aba_esclarecimento.xlsx"
    
    # Ports to use (make sure Chrome is running with these debugging ports)
    # Start Chrome with: chrome.exe --remote-debugging-port=9222 --user-data-dir=C:\selenium\ChromeProfile1
    portas = [9222, 9224, 9226, 9228]  # Add your ports here
    
    # Optional: Reset existing files
    reset = input('Do you want to reset all files? (y/n): ')
    if reset.lower() == 'y':
        for porta in portas:
            path = arquivo_saida_base.replace('.xlsx', f'_{porta}.xlsx')
            if os.path.exists(path):
                os.remove(path)
                print(f"🗑️ File deleted: {path}")
            else:
                print(f"⚠️ File not found: {path}")
    
    # Load proposals from Excel
    df = AsyncPWRobo.extrair_dados_excel(arquivo_fonte, filter_=True)
    if df is None:
        print("❌ Failed to load Excel file")
        return

    proposals_list = df.iloc[:, 0].tolist()  # Get first column as list
    proposals_list = [p for p in proposals_list if p and AsyncPWRobo.fix_prop_num(p)]
    print(f"\n📊 Total proposals to process: {len(proposals_list)}")

    # Collect all already processed proposal numbers from every port output file
    all_processed_numbers = set()
    for porta in portas:
        output_file = arquivo_saida_base.replace('.xlsx', f'_{porta}.xlsx')
        if os.path.exists(output_file):
            print(f"🔍 Checking existing file for port {porta}: {os.path.basename(output_file)}")
            try:
                df_existing = pd.read_excel(output_file)
                if 'Número da Proposta' in df_existing.columns and not df_existing.empty:
                    normalized = {
                        AsyncPWRobo.fix_prop_num(v)
                        for v in df_existing['Número da Proposta'].dropna().tolist()
                        if AsyncPWRobo.fix_prop_num(v)
                    }
                    all_processed_numbers.update(normalized)
                    print(f"✅ Port {porta}: found {len(normalized)} processed proposals")
                else:
                    print(f"⚠️ Port {porta}: Output file missing required columns or is empty")
            except Exception as e:
                print(f"⚠️ Port {porta}: Error reading existing file: {e}")

    remaining_proposals = [p for p in proposals_list if AsyncPWRobo.fix_prop_num(p) not in all_processed_numbers]
    print(f"🔁 Remaining proposals after filtering processed: {len(remaining_proposals)}")

    # Split remaining proposals into chunks for each port using stable ordering
    chunk_size = math.ceil(len(remaining_proposals) / len(portas)) if remaining_proposals else 0
    chunks = []

    for i, porta in enumerate(portas):
        start_idx = i * chunk_size
        end_idx = min(start_idx + chunk_size, len(remaining_proposals))
        chunk = remaining_proposals[start_idx:end_idx]
        chunks.append(chunk)
        print(f"📦 Port {porta}: Assigned {len(chunk)} proposals to process")
    
    # Create manager and process
    manager = AsyncScraperManager(portas, arquivo_saida_base)
    
    try:
        # Initialize all scrapers
        await manager.initialize_all()
        
        # Process each chunk on its assigned port
        tasks = []
        for robo, chunk in zip(manager.robos, chunks):
            if chunk:
                # Fix proposal numbers
                fixed_chunk = []
                for prop in chunk:
                    if prop in to_jump:
                        continue
                    fixed = AsyncPWRobo.fix_prop_num(prop)
                    if fixed:
                        fixed_chunk.append(fixed)
                    else:
                        print(f"⚠️ Could not fix proposal number: {prop}")
                
                if fixed_chunk:
                    print(f"\n🎯 Port {robo.porta} assigned {len(fixed_chunk)} proposals")
                    task = asyncio.create_task(manager._process_chunk(robo, fixed_chunk))
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
        print(f"✅ Successful instances: {successful}/{len(portas)}")
        print(f"❌ Failed instances: {failed}/{len(portas)}")
        
    except KeyboardInterrupt:
        print("\n🛑 Script stopped by user")
    except Exception as e:
        print(f"❌ Fatal error: {e}")
    finally:
        # Cleanup
        await manager.cleanup_all()


def main(to_jump: Optional[List[str]] = None):
    """Entry point - runs the async main"""
    asyncio.run(main_async(to_jump))


if __name__ == "__main__":
    to_jump = ['048032/2025', '060585/2025', '029262/2025']
    main(to_jump)