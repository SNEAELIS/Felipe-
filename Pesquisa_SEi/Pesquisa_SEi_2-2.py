import shutil
import re
import time
import os
import asyncio
from typing import List, Set, Dict

import pandas as pd

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError


class AsyncSEIScraper:
    def __init__(self, porta: int = 9222, arquivo_destino: str = None):
        self.porta = porta
        self.arquivo_destino = arquivo_destino
        self.playwright = None
        self.browser = None
        self.context = None
        self.page = None
        self.processados_no_lote = 0
        self.batch_size = 10
        self.results_buffer = []
        self.buffer_lock = asyncio.Lock()
        self.save_queue = asyncio.Queue()
        self.save_task = None

    async def conectar(self) -> bool:
        """Connect to existing Chrome browser via Playwright CDP"""
        try:
            print(f"🔌 Port {self.porta}: Connecting to browser via Playwright...")
            self.playwright = await async_playwright().start()
            cdp_url = f"http://127.0.0.1:{self.porta}"
            self.browser = await self.playwright.chromium.connect_over_cdp(cdp_url)
            self.context = self.browser.contexts[0] if self.browser.contexts else await self.browser.new_context()
            self.page = self.context.pages[0] if self.context.pages else await self.context.new_page()

            if await self.switch_to_sei():
                print(f"✅ Port {self.porta}: Connected successfully at {str(self.page.url)[:80]}")
                return True
            return False

        except Exception as e:
            print(f"❌ Port {self.porta}: Failed to connect: {type(e).__name__} - {e}")
            return False

    async def switch_to_sei(self) -> bool:
        """Find or switch to the SEI tab/page"""
        target_url = "sei.mds.gov.br"
        try:
            if self.page and target_url in self.page.url:
                print(f"✅ Port {self.porta}: Already on correct page")
                return True

            for context in self.browser.contexts:
                for page in context.pages:
                    if target_url in page.url:
                        self.page = page
                        print(f"🎯 Port {self.porta}: Switched to SEI page")
                        return True

            # Fallback: use first page if no exact match
            if self.context.pages:
                self.page = self.context.pages[0]
                print(f"⚠️ Port {self.porta}: No SEI page found, using first page: {self.page.url}")
                return target_url in self.page.url

            return False
        except Exception as e:
            print(f"❌ Port {self.porta}: SEI page not found: {type(e).__name__} - {e}")
            return False

    async def iniciar(self):
        """Initialize scraper and start save worker"""
        connected = await self.conectar()

        if not connected:
            raise Exception(f"Port {self.porta}: Failed to connect")

        # Start save worker
        self.save_task = asyncio.create_task(self._save_worker())

        # Load existing processed processes
        self.processados = await self._carregar_processados()

        return self

    async def _carregar_processados(self) -> Set[str]:
        """Load already processed processes from output file"""
        if not self.arquivo_destino or not os.path.exists(self.arquivo_destino):
            return set()

        try:
            loop = asyncio.get_event_loop()
            df = await loop.run_in_executor(None, lambda: pd.read_excel(self.arquivo_destino))

            if 'processo' in df.columns:
                processed = set(df['processo'].dropna().unique())
                print(f"📊 Port {self.porta}: Found {len(processed)} already processed")
                return processed
            return set()
        except Exception as e:
            print(f"⚠️ Port {self.porta}: Error loading processed: {e}")
            return set()

    async def _save_worker(self):
        """Background worker for saving data"""
        while True:
            try:
                # Wait for save signal or check every 2 seconds
                try:
                    signal = await asyncio.wait_for(self.save_queue.get(), timeout=2.0)
                    if signal == "STOP":
                        break
                except asyncio.TimeoutError:
                    # Timeout - check if we should save based on count
                    pass

                # Check if we need to save
                async with self.buffer_lock:
                    if len(self.results_buffer) >= self.batch_size:
                        await self._salvar_lote()
                    elif self.processados_no_lote > 0 and len(self.results_buffer) > 0:
                        # Save on STOP signal even if less than batch_size
                        if self.save_queue.empty() and self.processados_no_lote > 0:
                            await self._salvar_lote()

            except Exception as e:
                print(f"❌ Port {self.porta}: Save worker error: {e}")

    async def _salvar_lote(self):
        """Save current batch to Excel with timeout and fallback"""
        try:
            async with self.buffer_lock:
                if not self.results_buffer:
                    return

                df_batch = pd.DataFrame(self.results_buffer)
                # Clear buffer before save to allow new data while saving
                self.results_buffer = []
                self.processados_no_lote = 0

            # Apply filtering
            phrase = "Processo não possui andamentos abertos."
            df_cleaned = df_batch[
                (df_batch['texto_link'].str.contains('/', na=False) &
                 ~df_batch['texto_link'].str.contains(r'\.', na=False)) |
                (df_batch['texto_link'] == phrase)
                ]

            if df_cleaned.empty:
                print(f"💤 Port {self.porta}: No data to save in this batch")
                return

            print(f"💾 Port {self.porta}: Attempting to save {len(df_cleaned)} records...")

            # Save with timeout
            loop = asyncio.get_event_loop()
            try:
                # Run save with 30-second timeout
                await asyncio.wait_for(
                    loop.run_in_executor(
                        None,
                        lambda: self._append_to_excel_safe(df_cleaned, self.arquivo_destino)
                    ),
                    timeout=30.0
                )
                print(f"✅ Port {self.porta}: Successfully saved {len(df_cleaned)} records")

            except asyncio.TimeoutError:
                print(f"⏰ Port {self.porta}: Save operation timed out after 30s")
                # Try fallback: save to a temp file
                temp_file = self.arquivo_destino.replace('.xlsx', f'_temp_{int(time.time())}.xlsx')
                print(f"🔄 Port {self.porta}: Attempting fallback save to {os.path.basename(temp_file)}")
                try:
                    await loop.run_in_executor(
                        None,
                        lambda: df_cleaned.to_excel(temp_file, index=False)
                    )
                    print(f"✅ Port {self.porta}: Fallback save successful")
                    # Optionally notify user
                except Exception as e:
                    print(f"❌ Port {self.porta}: Fallback save also failed: {e}")

            except Exception as e:
                print(f"❌ Port {self.porta}: Save error: {type(e).__name__} - {e}")

        except Exception as e:
            print(f"❌ Port {self.porta}: Critical error in _salvar_lote: {e}")

    def _append_to_excel_safe(self, df_new_data: pd.DataFrame, arquivo_destino: str) -> bool:
        """Safely append to Excel (synchronous)"""
        try:
            # Read existing data
            if os.path.exists(arquivo_destino):
                df_existing = pd.read_excel(arquivo_destino)
            else:
                df_existing = pd.DataFrame()

            # Combine data
            if not df_existing.empty:
                # Align columns
                df_new_data = df_new_data.reindex(columns=df_existing.columns, fill_value='')
                df_combined = pd.concat([df_existing, df_new_data], ignore_index=True)
            else:
                df_combined = df_new_data

            # Remove duplicates
            df_combined = df_combined.drop_duplicates()

            # Save
            df_combined.to_excel(arquivo_destino, index=False)
            return True

        except Exception as e:
            print(f"❌ Error in append_to_excel_safe: {e}")
            return False

    async def extrair_processo(self, numero_processo: str) -> List[Dict]:
        """Extract links from a single process"""
        resultado = await self._extrair_links_async(numero_processo)

        # Add to buffer
        async with self.buffer_lock:
            self.results_buffer.extend(resultado)
            self.processados_no_lote += len(resultado)

            # Signal save if batch size reached
            if len(self.results_buffer) >= self.batch_size:
                await self.save_queue.put("SAVE")

        return resultado

    async def _extrair_links_async(self, numero_processo: str) -> List[Dict]:
        """Async extraction function using Playwright"""
        in_case_empty = [{
            "processo": numero_processo,
            "texto_link": "Erro: Sessão expirada ou acesso negado",
            "SNEALIS": "",
            "CGEALIS": "",
            "CGFP": "",
            "CGC": "",
            "CGAP": ""
        }]

        try:
            print(f"🔄 Port {self.porta}: Starting extraction for {numero_processo}")

            # Step 1: Find search field
            await self.page.wait_for_selector("#txtPesquisaRapida", timeout=7000)
            campo_busca = self.page.locator("#txtPesquisaRapida")
            print(f"📝 Port {self.porta}: Found search field")

            # Step 2: Enter process number
            print(f"✏️ Port {self.porta}: Entering process number: {numero_processo}")
            await campo_busca.fill(numero_processo)
            await campo_busca.press("Enter")
            await self.page.wait_for_load_state("networkidle", timeout=15000)

            # Step 3: Check for "no results"
            try:
                await self.page.wait_for_selector(".pesquisaSemResultado", timeout=1000)
                in_case_empty[0]["texto_link"] = "Nenhum resultado encontrado"
                return in_case_empty
            except PlaywrightTimeoutError:
                pass
            print(f"🔍 Port {self.porta}: Checked for no results")

            # Step 4: Locate tree content inside iframes
            div_arvore = await self._find_div_arvore_locator()
            await div_arvore.wait_for(state="visible", timeout=10000)
            print(f"📄 Port {self.porta}: Found divArvoreInformacao")

            # Step 5: Get raw text
            raw_text_total = (await div_arvore.inner_text()).strip()
            print(f"📝 Port {self.porta}: Got raw text")

            # Step 6: Find links
            link_handles = await div_arvore.locator("a").element_handles()
            print(f"🔗 Port {self.porta}: Found {len(link_handles)} link elements")

            resultados = []
            for el in link_handles:
                texto = (await el.inner_text()).strip()
                descricao = (await el.get_attribute("title")) or ""

                if texto:
                    texto_para_busca = f"{texto} {descricao}".upper()
                    snealis = "SNEALIS" if "SNEALIS" in texto_para_busca else ""
                    cgealis = "CGEALIS" if "CGEALIS" in texto_para_busca else ""
                    cgfp = "CGFP" if "CGFP" in texto_para_busca else ""
                    cgc = "CGC" if "CGC" in texto_para_busca else ""
                    cgap = "CGAP" if "CGAP" in texto_para_busca else ""

                    resultados.append({
                        "processo": numero_processo,
                        "texto_link": f"{texto} ({descricao})".strip(),
                        "SNEALIS": snealis,
                        "CGEALIS": cgealis,
                        "CGFP": cgfp,
                        "CGC": cgc,
                        "CGAP": cgap
                    })

            if not resultados and raw_text_total:
                print(f"⚠️ Port {self.porta}: No links found, using raw text as fallback")
                raw_upper = raw_text_total.upper()
                snealis = "SNEALIS" if "SNEALIS" in raw_upper else ""
                cgealis = "CGEALIS" if "CGEALIS" in raw_upper else ""
                cgfp = "CGFP" if "CGFP" in raw_upper else ""
                cgc = "CGC" if "CGC" in raw_upper else ""
                cgap = "CGAP" if "CGAP" in raw_upper else ""
                resultados.append({
                    "processo": numero_processo,
                    "texto_link": raw_text_total,
                    "SNEALIS": snealis,
                    "CGEALIS": cgealis,
                    "CGFP": cgfp,
                    "CGC": cgc,
                    "CGAP": cgap
                })

            print(f"✅ Port {self.porta}: Finished extraction for {numero_processo}")
            return resultados

        except Exception as e:
            print(f"❌ Port {self.porta}: ERROR processing {numero_processo}")
            print(f"  Error type: {type(e).__name__}")
            print(f"  Error message: {str(e)}")
            import traceback
            traceback.print_exc()
            return [{
                "processo": numero_processo,
                "texto_link": f"Erro: {type(e).__name__} - {str(e)[:50]}",
                "SNEALIS": "",
                "CGEALIS": "",
                "CGFP": "",
                "CGC": "",
                "CGAP": ""
            }]

    async def _find_div_arvore_locator(self):
        """Try nested iframe locators for divArvoreInformacao"""
        root_locator = self.page.locator("#divArvoreInformacao")
        if await root_locator.count() > 0:
            return root_locator

        nested_locator = self.page.frame_locator("iframe#ifrConteudoVisualizacao").frame_locator("iframe#ifrVisualizacao").locator("#divArvoreInformacao")
        if await nested_locator.count() > 0:
            return nested_locator

        direct_locator = self.page.frame_locator("iframe#ifrVisualizacao").locator("#divArvoreInformacao")
        if await direct_locator.count() > 0:
            return direct_locator

        # Try every frame in the page
        for frame in self.page.frames:
            frame_locator = frame.locator("#divArvoreInformacao")
            if await frame_locator.count() > 0:
                return frame_locator

        return root_locator

    async def cleanup(self):
        """Clean up resources"""
        if self.save_task:
            await self.save_queue.put("STOP")
            await self.save_task

        # Final save
        async with self.buffer_lock:
            if self.results_buffer:
                await self._salvar_lote()

        if self.playwright:
            try:
                await self.playwright.stop()
            except Exception:
                pass


class AsyncSEIManager:
    def __init__(self, portas: List[int], arquivo_fonte: str, arquivo_saida_base: str):
        self.portas = portas
        self.arquivo_fonte = arquivo_fonte
        self.arquivo_saida_base = arquivo_saida_base
        self.scrapers: List[AsyncSEIScraper] = []

    async def prepare_chunks_with_filter(self, df_completo: pd.DataFrame) -> List[List[str]]:
        """Prepare filtered chunks for each port"""
        chunks = []

        for i, porta in enumerate(self.portas):
            output_file = self.arquivo_saida_base.replace('.xlsx', f'_{porta}.xlsx')

            # Calculate chunk boundaries
            chunk_size = len(df_completo) // len(self.portas)
            start_idx = i * chunk_size
            end_idx = start_idx + chunk_size if i < len(self.portas) - 1 else len(df_completo)

            # Get raw chunk
            chunk = df_completo.iloc[start_idx:end_idx].copy()
            original_size = len(chunk)

            # Filter against existing output file
            if os.path.exists(output_file):
                try:
                    df_existing = pd.read_excel(output_file)
                    if 'processo' in df_existing.columns:
                        processed = set(df_existing['processo'].dropna().unique())
                        chunk = chunk[~chunk['processo'].isin(processed)]
                        print(f"📊 Port {porta}: Filtered out {original_size - len(chunk)} already processed")
                except Exception as e:
                    print(f"⚠️ Port {porta}: Error reading output file: {e}")

            chunks.append(chunk['processo'].tolist())
            print(f"📦 Port {porta}: {len(chunks[-1])} processes to process")

        return chunks

    async def initialize_all(self):
        """Initialize all scraper instances"""
        for i, porta in enumerate(self.portas):
            print(f"\n{'='*50}")
            print(f"Initializing port {porta} ({i+1}/{len(self.portas)})")

            output_file = self.arquivo_saida_base.replace('.xlsx', f'_{porta}.xlsx')
            scraper = AsyncSEIScraper(porta=porta, arquivo_destino=output_file)

            try:
                await scraper.iniciar()
                self.scrapers.append(scraper)
                print(f"✅ Port {porta} ready")
            except Exception as e:
                print(f"❌ Port {porta} failed: {e}")

            # Stagger initialization
            if i < len(self.portas) - 1:
                await asyncio.sleep(2)

    async def process_all(self, chunks: List[List[str]]):
        """Process all chunks in parallel"""
        tasks = []
        for scraper, chunk in zip(self.scrapers, chunks):
            if chunk:
                task = asyncio.create_task(self._process_chunk(scraper, chunk))
                tasks.append(task)

        return await asyncio.gather(*tasks, return_exceptions=True)

    async def _process_chunk(self, scraper: AsyncSEIScraper, processos: List[str]):
        """Process a chunk of processes with progress bar"""
        try:
            total = len(processos)

            # Process only first 3 for debugging
            for processo in processos[:3]:
                try:
                    await asyncio.wait_for(scraper.extrair_processo(processo), timeout=60.0)
                except asyncio.TimeoutError:
                    print(f"⏰ Port {scraper.porta}: Process {processo} timed out")
                except Exception as e:
                    print(f"❌ Port {scraper.porta}: Error processing {processo}: {e}")

            print(f"\n✅ Port {scraper.porta}: Completed 3 processes (debug mode)")

        except Exception as e:
            print(f"❌ Port {scraper.porta}: Error: {e}")
        finally:
            await scraper.cleanup()

    async def cleanup_all(self):
        """Clean up all instances"""
        for scraper in self.scrapers:
            await scraper.cleanup()


# --- Helper Functions (keep as is) ---
def formato_padrao(num_sei: str) -> str:
    """Transform SEI number to standard format"""
    num_sei = str(num_sei).strip()

    padrao = re.compile(r"^\d{5}\.\d{6}\/\d{4}-\d{2}$")
    if padrao.match(num_sei):
        return num_sei

    num_sei_limpo = re.sub(r'[./-]', '', num_sei)

    if len(num_sei_limpo) < 17:
        return ''

    part1 = num_sei_limpo[:5]
    part2 = num_sei_limpo[5:11]
    part3 = num_sei_limpo[11:15]
    part4 = num_sei_limpo[15:17]

    return f"{part1}.{part2}/{part3}-{part4}"


def ler_processos_validos(caminho_excel: str, nome_coluna: str = "processo") -> pd.DataFrame:
    """Read valid processes from Excel"""
    df = pd.read_excel(caminho_excel, dtype=str)
    print(f'📊 Loaded {len(df)} rows from Excel')

    # Clean and format processes
    df['processo'] = df[nome_coluna].apply(formato_padrao)

    # Filter valid format
    padrao = re.compile(r"^\d{5}\.\d{6}\/\d{4}-\d{2}$")
    df_valid = df[df['processo'].apply(lambda x: bool(padrao.match(str(x))))].copy()

    # Remove duplicates
    df_valid = df_valid.drop_duplicates(subset=['processo'])

    print(f"✅ {len(df_valid)} valid processes found")
    return df_valid[['processo']]


def delete_destiny_data(arquivo_destino, create_backup=True):
    """Delete all data from destiny file"""
    headers = ['processo', 'texto_link', 'SNEALIS', 'CGEALIS', 'CGFP', 'CGC', 'CGAP']

    if create_backup and os.path.exists(arquivo_destino):
        backup_path = arquivo_destino.replace('.xlsx', f'_backup.xlsx')
        shutil.copy2(arquivo_destino, backup_path)
        print(f"💾 Backup created: {backup_path}")

    empty_df = pd.DataFrame(columns=headers)
    empty_df.to_excel(arquivo_destino, index=False)
    print(f"🗑️ All data deleted from {arquivo_destino}")
    return True


# --- Main Async Function ---
async def main_async():
    """Main async function"""
    arquivo_fonte = (r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência "
                     r"Social\Teste001\propostas_SEi.xlsx")

    arquivo_saida_base = (r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência "
                          r"Social\SNEAELIS - webscraping\Consulta_SEi\Consultas parciais\consulta_direcao_sei.xlsx")

    # Ports to use
    portas = [9222, 9224, 9226, 9228]

    # Ask for reset
    #reset_df = input('Deseja resetar os DataFrames? [Y/n]: ')
    reset_df = 'n'
    if reset_df == 'Y':
        for porta in portas:
            output_file = arquivo_saida_base.replace('.xlsx', f'_{porta}.xlsx')
            delete_destiny_data(output_file)

    # Load processes
    df_processos = ler_processos_validos(arquivo_fonte)

    # Create manager
    manager = AsyncSEIManager(portas, arquivo_fonte, arquivo_saida_base)

    try:
        # Prepare filtered chunks
        chunks = await manager.prepare_chunks_with_filter(df_processos)

        # Initialize all scrapers
        await manager.initialize_all()

        if not manager.scrapers:
            print("❌ No scrapers initialized successfully")
            return

        # Process all chunks
        start_time = time.time()
        results = await manager.process_all(chunks)

        # Calculate statistics
        elapsed = time.time() - start_time
        successful = sum(1 for r in results if not isinstance(r, Exception))
        failed = sum(1 for r in results if isinstance(r, Exception))

        print(f"\n{'='*60}")
        print(f"🎉 ALL INSTANCES COMPLETED!")
        print(f"⏱️  Total time: {int(elapsed // 60)}m {int(elapsed % 60)}s")
        print(f"✅ Successful: {successful}/{len(portas)} instances")
        print(f"❌ Failed: {failed}/{len(portas)} instances")

    except KeyboardInterrupt:
        print("\n🛑 Stopped by user")
    finally:
        await manager.cleanup_all()


# --- Entry Point ---
if __name__ == "__main__":
    asyncio.run(main_async())