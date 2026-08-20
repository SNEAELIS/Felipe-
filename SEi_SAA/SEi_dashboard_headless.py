import shutil
import re
import sys
import time
import os
import traceback
import json
from pathlib import Path
from datetime import datetime

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError

os.system('cls' if os.name == 'nt' else 'clear')

# --- Garante que o navegador esteja na aba correta e conectado ao sei ---
def switch_to_sei(page):
    """Switch to the correct SEI page/tab"""
    target_url = "sei.mds.gov"
    
    # Check if we are already on the right page
    try:
        if target_url in page.url:
            print("✅ Already on the correct page.")
            return True
    except Exception:
        pass
    
    # Get all pages/tabs
    context = page.context
    all_pages = context.pages
    
    # Check all pages
    for p in all_pages:
        if target_url in p.url:
            print(f"🎯 Found and switched to: {p.url}")
            return p
    
    print("❌ Target URL not found in any open tabs.")
    return None


# --- Carrega credenciais de login do arquivo ---
def carregar_credenciais(arquivo_credenciais: str | Path = None):
    """Load SEI credentials from a JSON file."""
    if arquivo_credenciais is None:
        arquivo_credenciais = Path(__file__).parent / 'SEI_credentials.json'
    else:
        arquivo_credenciais = Path(arquivo_credenciais)

    if not arquivo_credenciais.exists():
        print(f"❌ Arquivo de credenciais não encontrado: {arquivo_credenciais}")
        print("📌 Crie um arquivo JSON contendo {\"username\": ..., \"password\": ..., \"login_url\": ...}")
        return None

    try:
        with arquivo_credenciais.open('r', encoding='utf-8') as f:
            dados = json.load(f)

        username = dados.get('username') or dados.get('usuario') or dados.get('login')
        password = dados.get('password') or dados.get('senha') or dados.get('pass')
        login_url = dados.get('login_url') or 'https://sei.mds.gov.br/sip/login.php?sigla_orgao_sistema=MC&sigla_sistema=SEI'

        if not username or not password:
            print('❌ Arquivo de credenciais inválido. Verifique username e password.')
            return None

        return {'username': username, 'password': password, 'login_url': login_url}
    except json.JSONDecodeError as e:
        print('❌ Erro ao ler credenciais JSON:', str(e))
        return None
    except Exception as e:
        print('❌ Erro inesperado ao carregar credenciais:', type(e).__name__, str(e)[:100])
        return None


# --- Inicia um navegador headless ---
def iniciar_navegador_headless(headless: bool = True):
    """Launch Playwright Chromium in headless mode."""
    try:
        playwright = sync_playwright().start()
        browser = playwright.chromium.launch(
            headless=headless,
            args=[
                '--disable-gpu',
                '--no-sandbox',
                '--disable-dev-shm-usage',
                '--disable-extensions',
                '--disable-infobars',
            ],
        )
        context = browser.new_context()
        page = context.new_page()
        return page, playwright, browser
    except Exception as e:
        print('❌ Erro ao iniciar navegador headless:', type(e).__name__, str(e)[:100])
        return None, None, None


# --- Verifica se a página atual é a tela de login do SEI ---
def is_login_page(page):
    try:
        login_url_fragments = ['/sip/login.php', '/sip/login', 'login']
        if any(fragment in page.url.lower() for fragment in login_url_fragments):
            return True

        selectors = [
            'input[name="txtLogin"]',
            'input[name="login"]',
            'input[name="usuario"]',
            'input[type="password"]',
            'input[name="txtSenha"]',
        ]

        for selector in selectors:
            if page.query_selector(selector):
                return True
    except Exception:
        pass
    return False


# --- Realiza login no SEI ---
def realizar_login_sei(page, credenciais: dict):
    """Fill the SEI login form and submit."""
    try:
        print('🔐 Abrindo a página de login do SEI...')
        page.goto(credenciais['login_url'], wait_until='domcontentloaded')

        login_selectors = [
            'input[name="txtLogin"]',
            'input[name="login"]',
            'input[name="usuario"]',
            'input[type="text"]',
        ]
        password_selectors = [
            'input[name="txtSenha"]',
            'input[name="senha"]',
            'input[type="password"]',
        ]
        submit_selectors = [
            'button[type="submit"]',
            'input[type="submit"]',
            'button[name="btnOK"]',
            'input[name="btnOK"]',
        ]

        login_field = None
        for selector in login_selectors:
            try:
                login_field = page.wait_for_selector(selector, timeout=6000)
                print(f'✓ Found login field: {selector}')
                break
            except Exception:
                continue

        password_field = None
        for selector in password_selectors:
            try:
                password_field = page.wait_for_selector(selector, timeout=6000)
                print(f'✓ Found password field: {selector}')
                break
            except Exception:
                continue

        if not login_field or not password_field:
            print('❌ Não foi possível localizar os campos de login e/ou senha.')
            return False

        login_field.fill(credenciais['username'])
        page.wait_for_timeout(200)
        password_field.fill(credenciais['password'])
        page.wait_for_timeout(200)

        clicked = False
        for selector in submit_selectors:
            try:
                button = page.query_selector(selector)
                if button:
                    button.click()
                    clicked = True
                    print(f'✓ Clicked submit button: {selector}')
                    break
            except Exception:
                continue

        if not clicked:
            page.keyboard.press('Enter')
            print('✓ Pressed Enter to submit login form')

        try:
            page.wait_for_load_state('networkidle', timeout=15000)
        except Exception:
            pass

        page.wait_for_timeout(1000)

        if is_login_page(page):
            print('❌ Login falhou ou a página de login ainda está visível. Verifique as credenciais e o CAPTCHA.')
            return False

        print('✅ Login realizado com sucesso! URL atual:', page.url)
        return True

    except Exception as e:
        print('❌ Erro ao realizar login:', type(e).__name__, str(e)[:120])
        return False


# --- Conectar ao navegador existente ---
def conectar_navegador_existente(porta: int):
    """Connect to existing Chrome instance via Playwright"""
    try:
        print(f"Tentando conectar ao navegador na porta {porta}...")
        
        playwright = sync_playwright().start()
        
        browser = playwright.chromium.connect_over_cdp(f"http://127.0.0.1:{porta}")
        
        if browser.contexts:
            context = browser.contexts[0]
            if context.pages:
                page = context.pages[0]
            else:
                page = context.new_page()
        else:
            context = browser.new_context()
            page = context.new_page()
        
        sei_page = switch_to_sei(page)
        if sei_page:
            page = sei_page

        print("Current URL:", page.url)
        print("✅ Conectado ao navegador existente com sucesso!")
        return page, playwright, browser
        
    except Exception as e:
        msg = "Erro ao conectar. Verifique se o Chrome está aberto com depuração."
        print("❌", msg)
        print(f'{type(e).__name__}\n\n{str(e)[:100]}')
        return None, None, None


# --- Transforma os números de processo para o formato padrão ---
def formato_padrao(num_sei: str):
    num_sei = str(num_sei).strip()
    
    padrao = re.compile(r"^\d{5}\.\d{6}\/\d{4}-\d{2}$")
    match_ = padrao.match(num_sei)
    
    if match_:
        return num_sei
    
    num_sei_limpo = re.sub(r'[./-]', '', num_sei)
    
    if len(num_sei_limpo) < 17:
        return ''
    else:
        part1 = num_sei_limpo[:5]
        part2 = num_sei_limpo[5:11]
        part3 = num_sei_limpo[11:15]
        part4 = num_sei_limpo[15:17]
        return f"{part1}.{part2}/{part3}-{part4}"


# --- Acessa o bloco de assinatura ---
def acessa_bloco_ass(page):
    """Access the signature block menu"""
    print('Acessando o bloco de assinaturas')
    try:
        xpaths = [
            'xpath=/html/body/div[1]/nav/div/div[3]/div[1]/div[1]/a',
            'xpath=/html/body/div[1]/div/div[1]/div[1]/ul/li[4]/a',
            'xpath=/html/body/div[1]/div/div[1]/div[1]/ul/li[4]/ul/li[1]/a'
        ]
        
        for xpath in xpaths:
            try:
                element = page.wait_for_selector(xpath, timeout=2000)
                element.click()
                time.sleep(0.5)
                print(f"✓ Clicked on {xpath}")
            except:
                continue
                
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred: {str(e)[:100]}")
        print(f"Error type: {exc_type.__name__}")
        print(f"Line number: {exc_tb.tb_lineno}")


# --- Extrai os processos do bloco de assinatura ---
def extrair_dados_bloco(page, arquivo_excel: str):
    """Extract data from signature block"""
    try:
        print(f"🔄 Iniciando processo de scraping - Bloco")
        
        # Wait for table
        try:
            table = page.wait_for_selector('#tblBlocos tbody', timeout=7000)
            rows = table.query_selector_all('tr')
            
            resultados = []
            
            for row in rows:
                cells = row.query_selector_all('td')
                if len(cells) > 5:
                    linha_dados = {
                        'Número': cells[1].text_content().strip() if cells[1] else '',
                        'Sinalizações': cells[2].text_content().strip() if cells[2] else '',
                        'Atribuição': cells[3].text_content().strip() if cells[3] else '',
                        'Estado': cells[4].text_content().strip() if cells[4] else '',
                        'Geradora': cells[5].text_content().strip() if cells[5] else '',
                        'Disponibilização': cells[6].text_content().strip() if cells[6] else '',
                        'Grupo': cells[7].text_content().strip() if cells[7] else '',
                        'Descrição': cells[8].text_content().strip() if cells[8] else '',
                        'Ações': cells[9].text_content().strip() if cells[9] else ''
                    }
                    resultados.append(linha_dados)
            
            if resultados:
                salvar_linha_excel(buffer_dados=resultados, arquivo=arquivo_excel, sheet_name='Bloco')
                print(f"Total de {len(resultados)} linhas salvas no Excel")
            else:
                print("Nenhuma linha encontrada para salvar")
            
            return resultados
            
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred: {str(e)[:100]}")
            print(f"Error type: {exc_type.__name__}")
            print(f"Line number: {exc_tb.tb_lineno}")
            
    except Exception as e:
        exc_tb = sys.exc_info()
        print(f"❌Error occurred at line: {exc_tb.tb_lineno}")
        print(f"{type(e).__name__} - {str(e)[:100]}")
        return None


# --- Extrai os processos das propostas ---
def extrair_dados_propostas(page, arquivo_excel: str):
    """Extract data from proposals"""
    try:
        print(f"🔄 Iniciando processo de scraping - Propostas")

        # Wait for block table
        table_bloco = page.wait_for_selector('#tblBlocos tbody', timeout=7000)
        linhas_bloco = table_bloco.query_selector_all('tr')
        total_rows = len(linhas_bloco)

        resultados = []

        for idx in range(1, total_rows):
            try:
                table_bloco = page.wait_for_selector('#tblBlocos tbody', timeout=7000)
                linhas_bloco = table_bloco.query_selector_all('tr')
                if idx >= len(linhas_bloco):
                    print(f"⚠️ Row {idx} not available after refresh, skipping")
                    continue

                linha = linhas_bloco[idx]
                cells = linha.query_selector_all('td')
                if len(cells) < 2:
                    continue

                numero = cells[1].text_content().strip() if cells[1] else ''
                link = cells[1].query_selector('a')
                if not link:
                    print(f"⚠️ No link found for row {idx} ({numero}), skipping")
                    continue

                link.click()
                time.sleep(1)

            except PlaywrightError as e:
                print(f"⚠️ Stale row or DOM issue at row {idx}: {type(e).__name__} {str(e)[:120]}")
                try:
                    page.wait_for_selector('#tblBlocos tbody', timeout=7000)
                except Exception:
                    print("⚠️ Could not recover block table after stale row")
                    break
                continue
            except Exception as e:
                print(f"⚠️ Unexpected error reading row {idx}: {type(e).__name__} {str(e)[:120]}")
                continue

            # After navigation, re-evaluate page state
            try:
                sem_registro = page.query_selector("xpath=//div[@id='divInfraAreaTabela']/label[contains(text(), 'Nenhum registro encontrado')]")

                if sem_registro:
                    linha_dados = {
                        'Número': numero,
                        'Seq.': 'Nenhum registro encontrado.',
                        'Processo': 'Nenhum registro encontrado.',
                        'Documento': 'Nenhum registro encontrado.',
                        'Tipo': 'Nenhum registro encontrado.',
                        'Assinaturas': 'Nenhum registro encontrado.',
                        'Anotações': 'Nenhum registro encontrado.',
                        'Ações': 'Nenhum registro encontrado.',
                    }
                    resultados.append(linha_dados)
                    page.go_back()
                    page.wait_for_selector('#tblBlocos tbody', timeout=7000)
                    continue

                table_pros = page.wait_for_selector('#tblProtocolosBlocos tbody', timeout=5000)
                linhas_pros = table_pros.query_selector_all('tr')

                for __, l in enumerate(linhas_pros):
                    if __ == 0:
                        continue

                    cells_pros = l.query_selector_all('td, th')
                    linha_dados = {
                        'Número': numero,
                        'Seq.': cells_pros[1].text_content().strip() if len(cells_pros) > 1 else '',
                        'Processo': cells_pros[2].text_content().strip() if len(cells_pros) > 2 else '',
                        'Documento': cells_pros[3].text_content().strip() if len(cells_pros) > 3 else '',
                        'Tipo': cells_pros[4].text_content().strip() if len(cells_pros) > 4 else '',
                        'Assinaturas': cells_pros[5].text_content().strip() if len(cells_pros) > 5 else '',
                        'Anotações': cells_pros[6].text_content().strip() if len(cells_pros) > 6 else '',
                        'Ações': cells_pros[7].text_content().strip() if len(cells_pros) > 7 else '',
                    }
                    resultados.append(linha_dados)

                page.go_back()
                page.wait_for_selector('#tblBlocos tbody', timeout=7000)

            except PlaywrightError as e:
                print(f"⚠️ Error after clicking process {numero}: {type(e).__name__} {str(e)[:120]}")
                try:
                    page.go_back()
                    page.wait_for_selector('#tblBlocos tbody', timeout=7000)
                except Exception:
                    print("⚠️ Could not recover after failed process detail extraction")
                    break
                continue
            except Exception as e:
                print(f"⚠️ Unexpected error extracting details for {numero}: {type(e).__name__} {str(e)[:120]}")
                try:
                    page.go_back()
                    page.wait_for_selector('#tblBlocos tbody', timeout=7000)
                except Exception:
                    print("⚠️ Could not recover after exception in details extraction")
                    break
                continue

        if resultados:
            salvar_linha_excel(buffer_dados=resultados, arquivo=arquivo_excel, sheet_name='Processos')
            print(f"Total de {len(resultados)} linhas salvas no Excel")
        else:
            print("Nenhuma linha encontrada para salvar")

        return resultados

    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred: {str(e)[:100]}")
        print(f"Error type: {exc_type.__name__}")
        print(f"Line number: {exc_tb.tb_lineno}")
        return None


# --- Espera os frames serem carregados ---
def wait_for_frames_to_load(page, expected_frame_names=None, timeout=30000):
    """
    Wait for frames to be loaded and available in page.frames
    
    Args:
        page: Playwright page object
        expected_frame_names: List of frame names to wait for (e.g., ['ifrArvore', 'ifrConteudoVisualizacao'])
        timeout: Maximum time to wait in milliseconds
    
    Returns:
        List of available frames
    """
    try:
        print("\n⏳ Waiting for frames to load...")
        
        start_time = time.time()
        frames_loaded = False
        last_frame_count = 0
        
        while time.time() - start_time < timeout / 1000:
            try:
                frames = page.frames
                current_frame_count = len(frames)
                
                if current_frame_count != last_frame_count:
                    print(f"  Checking frames... Found {current_frame_count} frame(s)")
                    last_frame_count = current_frame_count
                
                # If we have expected frames, check if they're all there
                if expected_frame_names:
                    found_frames = []
                    for name in expected_frame_names:
                        try:
                            frame = page.frame(name=name)
                            if frame:
                                found_frames.append(name)
                        except Exception as e:
                            # Silently continue if frame check fails
                            continue
                    
                    if len(found_frames) == len(expected_frame_names):
                       #print(f"✅ All expected frames found: {found_frames}")
                        frames_loaded = True
                        break
                    else:
                        # Only print every 3 seconds to avoid spam
                        if int(time.time() - start_time) % 1000 == 0:
                            print(f"  Found {len(found_frames)}/{len(expected_frame_names)} frames: ###{found_frames}")
                else:
                    # If no expected frames, wait for at least the iframes in DOM
                    try:
                        iframes = page.query_selector_all('iframe')
                        if len(iframes) > 0 and len(frames) > 1:
                            #print(f"✅ Found {len(frames)} frames and {len(iframes)} iframes")
                            frames_loaded = True
                            break
                    except Exception as e:
                        # If query_selector_all fails, continue
                        pass
                
                time.sleep(1)
                
            except Exception as e:
                exc_type, exc_value, exc_tb = sys.exc_info()
                print(f"⚠️ Error checking frames: {str(e)[:100]}")
                print(f"Error type: {exc_type.__name__}")
                print(f"Line number: {exc_tb.tb_lineno}")
                time.sleep(1)
                continue
        
        # If frames didn't load, try to force them
        if not frames_loaded:
            try:
                print("⚠️ Frames may not be fully loaded. Attempting to force load...")
                
                # Try scrolling to trigger lazy loading
                try:
                    page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
                    time.sleep(2)
                    page.evaluate("window.scrollTo(0, 0);")
                    time.sleep(2)
                    print("🔄 Scrolled to trigger lazy loading")
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(f"⚠️ Scroll failed: {str(e)[:100]}")
                    print(f"Error type: {exc_type.__name__}")
                    print(f"Line number: {exc_tb.tb_lineno}")
                
                # Try to trigger frame loading by interacting with the page
                try:
                    page.evaluate("""
                        () => {
                            // Try to trigger any lazy loading
                            window.dispatchEvent(new Event('scroll'));
                            window.dispatchEvent(new Event('resize'));
                            
                            // Try to trigger any iframe loading
                            const iframes = document.querySelectorAll('iframe');
                            iframes.forEach(iframe => {
                                try {
                                    iframe.src = iframe.src;
                                } catch(e) {}
                            });
                            return iframes.length;
                        }
                    """)
                    print("🔄 Triggered events to force frame loading")
                    time.sleep(2)
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(f"⚠️ Event trigger failed: {str(e)[:100]}")
                    print(f"Error type: {exc_type.__name__}")
                    print(f"Line number: {exc_tb.tb_lineno}")
                
                # Check again after forcing
                try:
                    frames = page.frames
                    print(f"📊 After forcing: {len(frames)} frames found")
                    
                    if expected_frame_names:
                        found_frames = []
                        for name in expected_frame_names:
                            frame = page.frame(name=name)
                            if frame:
                                found_frames.append(name)
                        
                        if len(found_frames) == len(expected_frame_names):
                            print(f"✅ All expected frames found after forcing: {found_frames}")
                            frames_loaded = True
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(f"⚠️ Force check failed: {str(e)[:100]}")
                    print(f"Error type: {exc_type.__name__}")
                    print(f"Line number: {exc_tb.tb_lineno}")
                
            except Exception as e:
                exc_type, exc_value, exc_tb = sys.exc_info()
                print(f"⚠️ Force loading failed: {str(e)[:100]}")
                print(f"Error type: {exc_type.__name__}")
                print(f"Line number: {exc_tb.tb_lineno}")
        
        # Final check
        try:
            final_frames = page.frames
            #print(f"📊 Final frame count: {len(final_frames)}")
            
            if expected_frame_names:
                available = []
                missing = []
                for name in expected_frame_names:
                    frame = page.frame(name=name)
                    if frame:
                        available.append(name)
                    else:
                        missing.append(name)
                
                if missing:
                    print(f"⚠️ Missing frames: {missing}")
                    print(f"✅ Available frames: {available}")
                else:
                    print(f"✅ All expected frames available: {available}")
            
            return final_frames
            
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"❌ Final frame check failed: {str(e)[:100]}")
            print(f"Error type: {exc_type.__name__}")
            print(f"Line number: {exc_tb.tb_lineno}")
            
            # Return whatever we have
            try:
                return page.frames
            except:
                return []
    
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"❌ wait_for_frames_to_load failed: {str(e)[:100]}")
        print(f"Error type: {exc_type.__name__}")
        print(f"Line number: {exc_tb.tb_lineno}")
        sys.exit()


# --- Captura exclusivamente o iframe "child" onde está a tabela ---
def wait_for_ifr_visualizacao_frame(page, timeout=20000):
    """Wait for the ifrVisualizacao frame to appear and return it."""
    print("⏳ Waiting for ifrVisualizacao frame...")
    deadline = time.time() + timeout / 1000.0
    last_error = None

    while time.time() < deadline:
        try:
            # Direct frame by name or id
            frame = page.frame(name='ifrVisualizacao')
            if frame and not frame.is_detached():
                #print(f"✅ Found ifrVisualizacao frame by name: {frame.url}")
                return frame

            iframe_element = page.query_selector('iframe[name="ifrVisualizacao"], iframe#ifrVisualizacao')
            if iframe_element:
                frame = iframe_element.content_frame()
                if frame and not frame.is_detached():
                    #print(f"✅ Found ifrVisualizacao via iframe element: {frame.url}")
                    return frame

            parent_iframe = page.query_selector('iframe[name="ifrConteudoVisualizacao"], iframe#ifrConteudoVisualizacao')
            if parent_iframe:
                parent_frame = parent_iframe.content_frame()
                if parent_frame and not parent_frame.is_detached():
                    child_iframe = parent_frame.query_selector('iframe[name="ifrVisualizacao"], iframe#ifrVisualizacao')
                    if child_iframe:
                        frame = child_iframe.content_frame()
                        if frame and not frame.is_detached():
                            #print(f"✅ Found ifrVisualizacao inside parent frame: {frame.url}")
                            return frame

            for f in page.frames:
                if f.is_detached():
                    continue
                if f.name == 'ifrVisualizacao' or 'visualizacao' in str(f.name).lower() or 'arvore_processar_html' in f.url:
                   #print(f"✅ Found ifrVisualizacao by frame list: name={f.name} url={f.url}")
                    return f

        except Exception as e:
            last_error = e

        time.sleep(0.5)

    if last_error:
        print(f"⚠️ wait_for_ifr_visualizacao_frame timeout. Last error: {type(last_error).__name__}: {str(last_error)[:120]}")
    else:
        print("⚠️ wait_for_ifr_visualizacao_frame timeout without finding the frame.")
    return None


# --- Extrai os dados da tabela de histórico ---
def extract_history_from_nested_frame(page, numero_processo) -> list[dict]:
    """Extract history table from the nested ifrVisualizacao frame"""
    
    try:
        print(f"\n📋 Extracting history for: {numero_processo}")

        viz_frame = wait_for_ifr_visualizacao_frame(page, timeout=20000)
        if not viz_frame:
            print("❌ ifrVisualizacao frame not found")
            print("\nAll available frames:")
            for f in page.frames:
                #print(f"  - {f.name}: {f.url[:80]}")
                continue
            return None

        print("✅ Found ifrVisualizacao frame")
        #print(f"   URL: {viz_frame.url}")

        # Wait for content inside ifrVisualizacao
        try:
            viz_frame.wait_for_selector('table, .infraAreaTelad, #tblHistorico, div.infraAreaTelad, tbody', timeout=15000)
            print("✓ Content loaded in ifrVisualizacao")
            if viz_frame.is_detached():
                print("⚠️ ifrVisualizacao frame detached after load; reacquiring")
                viz_frame = wait_for_ifr_visualizacao_frame(page, timeout=10000)
                if not viz_frame:
                    return None
        except Exception as e:
            print(f"⚠️ Content may not be fully loaded: {type(e).__name__} {str(e)[:120]}")
            viz_frame = wait_for_ifr_visualizacao_frame(page, timeout=10000)
            if not viz_frame:
                return None

        # Step 6: Extract the table data
        try:
            js_source = """
            () => {
                const result = [];
                
                // Try to find the history table
                const tables = document.querySelectorAll('table');
                let historyTable = null;
                
                // Look for the table with history data
                for (let table of tables) {
                    const text = table.textContent;
                    if (text.includes('Data/Hora') || text.includes('Data') || 
                        text.includes('Andamentos') || text.includes('Histórico')) {
                        historyTable = table;
                        break;
                    }
                }
                
                // If no specific table found, use the first table with data
                if (!historyTable) {
                    for (let table of tables) {
                        const rows = table.querySelectorAll('tr');
                        if (rows.length > 1) {
                            const firstRow = rows[0];
                            const cells = firstRow.querySelectorAll('td, th');
                            if (cells.length >= 4) {
                                historyTable = table;
                                break;
                            }
                        }
                    }
                }
                
                if (historyTable) {
                    const rows = historyTable.querySelectorAll('tr');
                    for (let row of rows) {
                        const cells = row.querySelectorAll('td');
                        if (cells.length >= 4) {
                            const rowData = {
                                'Processo': PROCESSO_PLACEHOLDER,
                                'Data/Hora': cells[0].textContent.trim(),
                                'Unidade': cells[1].textContent.trim(),
                                'Usuário': cells[2].textContent.trim(),
                                'Descrição': cells[3].textContent.trim()
                            };
                            // Only add if there's actual data
                            if (rowData['Data/Hora'] || rowData['Descrição']) {
                                result.push(rowData);
                            }
                        }
                    }
                }
                
                // If no table found, try to get all text
                if (result.length === 0) {
                    const bodyText = document.body ? document.body.innerText : '';
                    if (bodyText) {
                        // Try to parse text as table
                        const lines = bodyText.split('\\n').filter(line => line.trim());
                        for (let line of lines) {
                            const parts = line.split(/\\s{2,}/);
                            if (parts.length >= 4) {
                                result.push({
                                    'Processo': PROCESSO_PLACEHOLDER,
                                    'Data/Hora': parts[0].trim(),
                                    'Unidade': parts[1] ? parts[1].trim() : '',
                                    'Usuário': parts[2] ? parts[2].trim() : '',
                                    'Descrição': parts.slice(3).join(' ').trim()
                                });
                            }
                        }
                    }
                }
                
                return result;
            }
        """
            results = viz_frame.evaluate(js_source.replace('PROCESSO_PLACEHOLDER', json.dumps(numero_processo)))
        except PlaywrightError as e:
            print(f"⚠️ Frame evaluation failed: {type(e).__name__} {str(e)[:120]}")
            viz_frame = wait_for_ifr_visualizacao_frame(page, timeout=10000)
            if not viz_frame:
                return None
            js_source = """
            () => {
                const result = [];
                
                // Try to find the history table
                const tables = document.querySelectorAll('table');
                let historyTable = null;
                
                // Look for the table with history data
                for (let table of tables) {
                    const text = table.textContent;
                    if (text.includes('Data/Hora') || text.includes('Data') || 
                        text.includes('Andamentos') || text.includes('Histórico')) {
                        historyTable = table;
                        break;
                    }
                }
                
                // If no specific table found, use the first table with data
                if (!historyTable) {
                    for (let table of tables) {
                        const rows = table.querySelectorAll('tr');
                        if (rows.length > 1) {
                            const firstRow = rows[0];
                            const cells = firstRow.querySelectorAll('td, th');
                            if (cells.length >= 4) {
                                historyTable = table;
                                break;
                            }
                        }
                    }
                }
                
                if (historyTable) {
                    const rows = historyTable.querySelectorAll('tr');
                    for (let row of rows) {
                        const cells = row.querySelectorAll('td');
                        if (cells.length >= 4) {
                            const rowData = {
                                'Processo': PROCESSO_PLACEHOLDER,
                                'Data/Hora': cells[0].textContent.trim(),
                                'Unidade': cells[1].textContent.trim(),
                                'Usuário': cells[2].textContent.trim(),
                                'Descrição': cells[3].textContent.trim()
                            };
                            if (rowData['Data/Hora'] || rowData['Descrição']) {
                                result.push(rowData);
                            }
                        }
                    }
                }
                
                if (result.length === 0) {
                    const bodyText = document.body ? document.body.innerText : '';
                    if (bodyText) {
                        const lines = bodyText.split('\n').filter(line => line.trim());
                        for (let line of lines) {
                            const parts = line.split(/\s{2,}/);
                            if (parts.length >= 4) {
                                result.push({
                                    'Processo': PROCESSO_PLACEHOLDER,
                                    'Data/Hora': parts[0].trim(),
                                    'Unidade': parts[1] ? parts[1].trim() : '',
                                    'Usuário': parts[2] ? parts[2].trim() : '',
                                    'Descrição': parts.slice(3).join(' ').trim()
                                });
                            }
                        }
                    }
                }
                
                return result;
            }
        """
            results = viz_frame.evaluate(js_source.replace('PROCESSO_PLACEHOLDER', json.dumps(numero_processo)))

        print(f"✅ Extracted {len(results)} history entries")
        
        if results:
            for entry in results[:3]:  # Show first 3 entries
                print(f"  {entry['Data/Hora']} | {entry['Unidade']} | {entry['Usuário']} | {entry['Descrição'][:50]}")
        
        return results
        
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"❌ Error extracting history: {str(e)[:100]}")
        print(f"Error type: {exc_type.__name__}")
        print(f"Line number: {exc_tb.tb_lineno}")
        return None


# --- Extrai detalhes das propostas (andamentos) ---
def get_proposal_details(page, arquivo_excel: str, sheet_name: str):
    """Extract proposal tracking details and save to specified sheet_name"""
    
    try:
        # Read Excel file
        if sheet_name == 'Andamento':
            df = pd.read_excel(arquivo_excel, dtype=str, sheet_name='Processos')
            pattern = r'^\d{5}\.\d{6}/\d{4}-\d{2}$'
            df = df[df['Processo'].str.match(pattern, na=False)]
            df = df.drop_duplicates(subset=['Processo'])
            numeros_processo = df['Processo'].tolist()
        else:
            page.click('#lnkControleProcessos')
            source, new_processes = extract_received_processos(page=page)
            numeros_processo = find_missing_processos(arquivo_excel,
                                                     source=source,
                                                     target_sheet='Processos')
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error reading Excel: {str(e)[:100]}")
        print(f"Error type: {exc_type.__name__}")
        print(f"Line number: {exc_tb.tb_lineno}")
        return
    
    all_results = []
    for numero_processo in numeros_processo:
        try:
            wait_for_frames_to_load
            print(f"\n📋 Processando: {numero_processo}")
            
            # Go to default content (top frame)
            # Playwright handles frames differently - we need to find the main frame
            
            # Find search field
            try:
                # First, try to find on main page
                campo_busca = page.wait_for_selector('#txtPesquisaRapida', timeout=5000)
                campo_busca.fill('')
                campo_busca.fill(numero_processo)
                campo_busca.press('Enter')
                time.sleep(0.75)
            except:
                # Try alternative selector
                campo_busca = page.wait_for_selector('input[name="txtPesquisaRapida"]', timeout=5000)
                campo_busca.fill('')
                campo_busca.fill(numero_processo)
                campo_busca.press('Enter')
                time.sleep(0.75)

            main_page = page 
            frames = wait_for_frames_to_load(page=page, expected_frame_names=None, timeout=30000) 
            
            # Find and switch to iframeArvore
            frame_arvore = page.frame(name='ifrArvore')
            if not frame_arvore:
                frame_arvore = page.frame(url='*arvore_visualizar*')
            
            if frame_arvore:
                # Click on "Consultar Andamento"
                try:
                    consultar_link = frame_arvore.wait_for_selector('#divConsultarAndamento a', timeout=5000)
                    consultar_link.click()
                    time.sleep(0.25)
                except:
                    print(f"⚠️ 'Consultar Andamento' not found for {numero_processo}")
                    continue
            
            # Find the main content frame
            page = main_page

            try:
                attempt = 0
                while attempt <= 3:
                    resultados = extract_history_from_nested_frame(page=page, numero_processo=numero_processo)
                    
                    if resultados:
                        all_results.extend(resultados)
                        print(f"✅ {len(resultados)} registros de andamento coletados para {numero_processo}")
                        break
                    else:
                        time.sleep(1)
                        attempt += 1
                        print(f"⚠️ Nenhum andamento encontrado para {numero_processo}.\nTentando novamento, tentativa número{attempt}")
                
            except Exception as e:
                print(f"⚠️ Error extracting history for {numero_processo}: {str(e)[:100]}")
                continue
                
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error processing {numero_processo}: {str(e)[:100]}")
            print(f"Error type: {exc_type.__name__}")
            print(f"Line number: {exc_tb.tb_lineno}")
            sys.exit()
    if not  sheet_name == 'Andamento':
        if new_processes:
            all_results.extend(new_processes)
            print(f"✅ Adicionados {len(new_processes)} processos não visualizados (vermelhos)")

    if all_results:
        salvar_linha_excel(buffer_dados=all_results, arquivo=arquivo_excel, sheet_name=sheet_name)
        print(f"✅ Total de {len(all_results)} registros salvos na planilha '{sheet_name}' do arquivo {arquivo_excel}")
    else:
        print(f"⚠️ Nenhum registro coletado. Nenhuma gravação realizada.")

# --- Extrai os dados de cada processo ---
def extract_received_processos(page) -> pd.DataFrame:
    """Extract all process numbers from the received processes table.
    
    Filters out red-colored text, keeping all other colors.
    """
    try:
        # Get all process numbers and filter out red ones in one go
        result = page.evaluate('''
            () => {
                const rows = document.querySelectorAll('#tblProcessosRecebidos tbody tr');
                const results = [];
                let skipped = [];
                
                for (let i = 1; i < rows.length; i++) {
                    const cells = rows[i].querySelectorAll('td');
                    
                    if (cells.length >= 3) {
                        const cell = cells[2];
                        const text = cell.textContent.trim();
                        
                        // Check if any element in the cell has the red class
                        const hasRedClass = cell.querySelector('.processoNaoVisualizado') !== null;
                        
                        // Also check if the cell itself has the class
                        const cellHasClass = cell.classList.contains('processoNaoVisualizado');
                        
                        if (hasRedClass || cellHasClass) {
                            // This is a red (unviewed) process - skip it
                            skipped.push(text);
                        } else if (text && text.includes('.')) {
                            // Valid process number (black/default color)
                            results.push(text);
                        } else {
                            skipped.push(text);
                        }
                    } else {
                        skipped.push(text);
                    }
                }
                
                return {
                    processos: results,
                    skipped: skipped
                };
            }
        ''')
        
        skipped_as_dicts = []
        for process in result['skipped']:
            skipped_as_dicts.append({
                'Processo': process,
                'Data/Hora': 'novo',
                'Unidade': 'novo',
                'Usuário': 'novo',
                'Descrição': 'novo'
                # Add any other fields that your history extraction returns
            })
        
        df = pd.DataFrame({'Processo': result['processos']})

        print(f"✅ Extracted {len(df)} process numbers (skipped {result['skipped']} unviewed entries)")
        return df, skipped_as_dicts

    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"❌ Error extracting received process numbers: {str(e)[:100]}")
        print(f"Error type: {exc_type.__name__} at line {exc_tb.tb_lineno}")
        return pd.DataFrame(columns=['Processo'])

# --- Filtra os processos que estão na planilha fonte porém não estão no planilha alvo ---
def find_missing_processos(arquivo_excel: str, source: str, target_sheet: str = 'Processos', filter: bool = False) -> list:
    """Return list of `Processo` values present in `source_sheet` but not in `target_sheet`.

    Args:
        arquivo_excel: path to the Excel file
        source_sheet: sheet name to compare from (contains candidate `Processo` values)
        target_sheet: sheet to compare against (defaults to 'Processos')

    Returns:
        List of Processo strings that are in source_sheet but missing from target_sheet.
    """
    try:
        # Read both sheets
        df_target = pd.read_excel(arquivo_excel, dtype=str, sheet_name=target_sheet)

        if 'Processo' not in source.columns:
            print(f"⚠️ 'Processo' column not found in source sheet '{source}'")
            return []
        if 'Processo' not in df_target.columns:
            print(f"⚠️ 'Processo' column not found in target sheet '{target_sheet}'")
            return []

        src_set = set(source['Processo'].dropna().astype(str).str.strip().unique())

        tgt_set = set(df_target['Processo'].dropna().astype(str).str.strip().unique())
        # Only use if filter is required
        if filter:
            missing = sorted(list(src_set - tgt_set))
        else:
            missing = sorted(list(src_set))

        print(f"🔎 Found {len(missing)} Processo(s)  missing from '{target_sheet}'")
        return missing

    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"❌ Error comparing sheets: {str(e)[:100]}")
        print(f"Error type: {exc_type.__name__} at line {exc_tb.tb_lineno}")
        return []

# --- Debugger para os Iframes ---
def debug_all_frames(page, single_frame:str=''):
    """Debug all frames in Playwright"""
    
    print("\n" + "="*60)
    print("DEBUGGING ALL FRAMES IN PLAYWRIGHT")
    print("="*60)
    
    # Method 1: Get frames from page.frames
    frames = page.frames
    print(f"\n📊 Total frames in page.frames: {len(frames)}")
    
    for idx, frame in enumerate(frames):
        if single_frame:
            if frame.name!=single_frame:
                continue
        print(f"\n--- FRAME {idx} ---")
        print(f"  Name: {frame.name}")
        print(f"  URL: {frame.url}")
        print(f"  Is main: {frame == page.main_frame}")
        
        # Try to get frame content
        try:
            title = frame.evaluate('document.title')
            print(f"  Title: {title}")
            
            body_text = frame.evaluate('document.body ? document.body.innerText.substring(0, 200) : ""')
            print(f"  Body text preview: {body_text[:100] if body_text else 'Empty'}")
            
            element_count = frame.evaluate('document.querySelectorAll("*").length')
            print(f"  Elements: {element_count}")
        except Exception as e:
            print(f"  ❌ Cannot access frame: {e}")
    
    # Method 2: Find iframes in the DOM
    print("\n" + "="*60)
    print("FINDING IFRAMES IN DOM")
    print("="*60)
    
    # Get all iframe elements from the main page
    iframe_elements = page.query_selector_all('iframe')
    print(f"\n📊 Iframe elements found in main page: {len(iframe_elements)}")
    
    for idx, iframe in enumerate(iframe_elements):
        print(f"\n--- IFRAME ELEMENT {idx} ---")
        
        # Get iframe attributes
        id_attr = iframe.get_attribute('id')
        name_attr = iframe.get_attribute('name')
        src_attr = iframe.get_attribute('src')
        class_attr = iframe.get_attribute('class')
        
        print(f"  ID: {id_attr}")
        print(f"  Name: {name_attr}")
        print(f"  SRC: {src_attr}")
        print(f"  Class: {class_attr}")
        
        # Try to get the content frame
        try:
            content_frame = iframe.content_frame()
            if content_frame:
                print(f"  ✅ Content frame found!")
                print(f"     Frame URL: {content_frame.url}")
                print(f"     Frame Name: {content_frame.name}")
            else:
                print(f"  ❌ No content frame (might be cross-origin or not loaded)")
        except Exception as e:
            print(f"  ❌ Error getting content frame: {e}")
    
    # Method 3: Check if frames are nested
    print("\n" + "="*60)
    print("CHECKING FOR NESTED FRAMES")
    print("="*60)
    
    for frame in page.frames:
        if frame != page.main_frame:
            try:
                # Check if this frame has its own iframes
                child_iframes = frame.query_selector_all('iframe')
                if child_iframes:
                    print(f"\nFrame {frame.name} has {len(child_iframes)} child iframes:")
                    for child in child_iframes:
                        print(f"  - ID: {child.get_attribute('id')}, Name: {child.get_attribute('name')}")
            except:
                pass
    
    return frames, iframe_elements


# --- Confere iframes ---
def confere_iframe(page, iframe_id: str, dbg: bool = False):
    """Check and switch to specified iframe"""
    try:
        page.wait_for_selector('iframe', timeout=7000)
    except:
        print("No iframes found on page")
        return
    
    if dbg:
        debug_all_frames(page)
    
    # Try to find and switch to the specific iframe
    try:
        # Look for iframe by ID
        iframe_element = page.query_selector(f'#{iframe_id}')
        if iframe_element:
            frame = page.frame(name=iframe_id)
            if frame:
                print(f"✓ Switched to frame: {iframe_id}")
                return frame
        
        # Try by URL pattern
        for frame in page.frames:
            if 'conteudo' in frame.url or 'procedimento' in frame.url:
                print(f"✓ Found content frame: {frame.url}")
                return frame
        
        print(f"⚠️ Could not find frame: {iframe_id}")
        return None
        
    except Exception as e:
        print(f"Error finding iframe: {type(e).__name__} - {str(e)[:80]}")
        return None


# --- Salva os dados no excel ---
def salvar_linha_excel(buffer_dados: list, arquivo: str | Path, sheet_name: str):
    """Save data to Excel file"""
    try:
        limpar_planilha(arquivo_fonte=arquivo, sheet_name=sheet_name)
        
        if sheet_name == 'Bloco':
            colunas = ['Número', 'Sinalizações', 'Atribuição', 'Estado', 
                       'Geradora', 'Disponibilização', 'Grupo', 'Descrição', 'Ações']
        elif sheet_name == 'Processos':
            colunas = ['Número', 'Seq.', 'Processo', 'Documento', 'Tipo', 
                       'Assinaturas', 'Anotações', 'Ações']
        elif sheet_name == 'Andamento':
            colunas = ['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição']
        elif sheet_name == 'Controle de Processo':
            colunas = ['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição']
        
        # Load existing sheets
        if os.path.exists(arquivo):
            with pd.ExcelFile(arquivo) as xlsx:
                sheets = {sheet: pd.read_excel(arquivo, sheet_name=sheet) 
                         for sheet in xlsx.sheet_names}
            
            if sheet_name in sheets:
                df_target = sheets[sheet_name]
            else:
                df_target = pd.DataFrame(columns=colunas)
        else:
            sheets = {}
            df_target = pd.DataFrame(columns=colunas)
        
        # Create new rows
        novas_linhas = []
        for linha in buffer_dados:
            nova_linha = {col: linha.get(col, '') for col in colunas}
            novas_linhas.append(nova_linha)
        
        # Combine
        df_final = pd.concat([df_target, pd.DataFrame(novas_linhas)], ignore_index=True)
        sheets[sheet_name] = df_final
        
        # Save
        with pd.ExcelWriter(arquivo, engine='openpyxl') as writer:
            for sheet, df in order_sheets(sheets):
                df.to_excel(writer, sheet_name=sheet, index=False)
        
        print(f"✅ {len(buffer_dados)} linhas salvas em '{sheet_name}' no arquivo {arquivo}")
        return True
        
    except Exception as e:
        print(f"❌ Erro ao salvar no Excel: {type(e).__name__}\n{str(e)[:100]}")
        return False

# --- Ordena as planilhas em sequência específica ---
def order_sheets(sheets: dict) -> list[tuple[str, object]]:
    """Return workbook sheets in the fixed desired order."""
    SHEET_ORDER = ['Bloco', 'Processos', 'Andamento', 'Controle de Processo']

    ordered = []
    for sheet_name in SHEET_ORDER:
        if sheet_name in sheets:
            ordered.append((sheet_name, sheets[sheet_name]))
    for sheet_name, sheet_df in sheets.items():
        if sheet_name not in SHEET_ORDER:
            ordered.append((sheet_name, sheet_df))
    return ordered

# --- Deleta todos os dados no arquivo de saída ---
def limpar_planilha(arquivo_fonte: str | Path, sheet_name: str):
    """Clear sheet data while preserving headers"""
    
    if sheet_name == 'Bloco':
        headers = ['Número', 'Sinalizações', 'Atribuição', 'Estado', 
                   'Geradora', 'Disponibilização', 'Grupo', 'Descrição', 'Ações']
    elif sheet_name == 'Processos':
        headers = ['Número', 'Seq.', 'Processo', 'Documento', 'Tipo', 
                   'Assinaturas', 'Anotações', 'Ações']
    elif sheet_name == 'Andamento':
        headers = ['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição']
    elif sheet_name == 'Controle de Processo':
        headers = ['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição']
    
    if not os.path.exists(arquivo_fonte):
        print(f"📁 File not found: {arquivo_fonte}")
        print(f"📝 Creating new file with sheet '{sheet_name}'...")
        try:
            empty_df = pd.DataFrame(columns=headers)
            empty_df.to_excel(arquivo_fonte, sheet_name=sheet_name, index=False)
            print(f"✅ File created with sheet '{sheet_name}'")
            return True
        except Exception as e:
            print(f"❌ Error creating file: {type(e).__name__}.\nMSG: {str(e)[:100]}")
            return False
    
    try:
        with pd.ExcelFile(arquivo_fonte) as xlsx:
            sheets = {sheet: pd.read_excel(arquivo_fonte, sheet_name=sheet) 
                     for sheet in xlsx.sheet_names}
        
        if sheet_name in sheets:
            empty_df = pd.DataFrame(columns=headers)
            sheets[sheet_name] = empty_df
            print(f"🗑️ All data deleted from sheet '{sheet_name}'")
        else:
            empty_df = pd.DataFrame(columns=headers)
            sheets[sheet_name] = empty_df
            print(f"📝 Sheet '{sheet_name}' not found. Creating new sheet with headers only.")
        
        with pd.ExcelWriter(arquivo_fonte, engine='openpyxl') as writer:
            for sheet, df in order_sheets(sheets):
                df.to_excel(writer, sheet_name=sheet, index=False)
        
        print(f"✅ Sheet '{sheet_name}' cleared successfully")
        return True
        
    except Exception as e:
        print(f"❌ Error clearing sheet: {type(e).__name__}.\nMSG: {str(e)[:100]}")
        return False

# --- Concatena os dados no arquivo excel ---
def append_to_excel_safe(df_new_data, arquivo_fonte, make_backup=True):
    """Safely append new data to existing Excel file"""
    try:
        if make_backup and os.path.exists(arquivo_fonte):
            backup_path = arquivo_fonte.replace('.xlsx', f'_backup.xlsx')
            shutil.copy2(arquivo_fonte, backup_path)
            print(f"💾 Backup created: {backup_path}")
        
        if os.path.exists(arquivo_fonte):
            try:
                df_existing = pd.read_excel(arquivo_fonte)
                print(f"📖 Read {len(df_existing)} existing rows from {arquivo_fonte}")
            except Exception as e:
                print(f"⚠️ Could not read existing file: {type(e).__name__}.\nMSG: {str(e)[:100]}")
                df_existing = pd.DataFrame()
        else:
            df_existing = pd.DataFrame()
            print(f"📁 File doesn't exist yet, will create new")
        
        if not isinstance(df_new_data, pd.DataFrame):
            try:
                df_new_data = pd.DataFrame(df_new_data)
                print(f"🔄 Converted new data to DataFrame")
            except:
                print(f"❌ Could not convert new data to DataFrame")
                return False
        
        if not df_existing.empty:
            if set(df_new_data.columns) != set(df_existing.columns):
                print(f"⚠️ Column mismatch. Aligning new data to existing columns...")
                df_new_data = df_new_data.reindex(columns=df_existing.columns)
            
            df_combined = pd.concat([df_existing, df_new_data], ignore_index=True)
            print(f"🔗 Combined {len(df_existing)} existing + {len(df_new_data)} new rows")
        else:
            df_combined = df_new_data
            print(f"📝 Using only new data ({len(df_new_data)} rows)")
        
        before_dedup = len(df_combined)
        df_combined = df_combined.drop_duplicates()
        after_dedup = len(df_combined)
        if before_dedup > after_dedup:
            print(f"🗑️ Removed {before_dedup - after_dedup} duplicate rows")
        
        df_combined.to_excel(arquivo_fonte, index=False)
        print(f"✅ Data saved to {arquivo_fonte}")
        print(f"   Total rows: {len(df_combined)} (Added {len(df_new_data)} new rows)")
        
        return True
        
    except Exception as e:
        print(f"❌ Error saving to Excel: {type(e).__name__}")
        print(f"   Message: {str(e)[:100]}")
        return False

# --- Função principal ---
def executar_scraping():
    """Main execution function"""
    
    arquivo_fonte = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\Consulta_SEi\dashboard_se\DB_sei_se.xlsx"
    arquivo_destino = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\Controle_SEI\DB_sei_se.xlsx"
    credenciais_arquivo = None
    
    page = None
    playwright = None
    browser = None
    
    try:
        credenciais = carregar_credenciais(credenciais_arquivo)
        if not credenciais:
            return

        page, playwright, browser = iniciar_navegador_headless(headless=True)
        if not page:
            print("❌ Não foi possível iniciar o navegador headless. Saindo.")
            return

        if not realizar_login_sei(page, credenciais):
            print("❌ Login não concluído. Verifique as credenciais e execute novamente.")
            return

        # Main loop
        acessa_bloco_ass(page=page)

        while True:
            try:
                extrair_dados_bloco(page=page,
                                     arquivo_excel=arquivo_fonte
                                     )
                
                extrair_dados_propostas(page=page,
                                        arquivo_excel=arquivo_fonte
                                        )
                
                get_proposal_details(page=page,
                                     arquivo_excel=arquivo_fonte,
                                     sheet_name='Andamento'
                                     )
                
                get_proposal_details(page=page,
                                     arquivo_excel=arquivo_fonte,
                                     sheet_name='Controle de Processo'
                                     )
                
                try:
                    page.click('#lnkControleProcessos')
                except Exception:
                    pass
                shutil.copy(arquivo_fonte, arquivo_destino)
                print(f"✅ Copied file to {arquivo_destino}")
                
                #break
                print("⏳ Waiting 600 seconds before next run...")
                time.sleep(600)
                
            except Exception as e:
                print(f"Error in main loop: {type(e).__name__}")
                print(f"Message: {str(e)[:100]}")
                traceback.print_exc()
                time.sleep(60)  # Wait a minute before retrying
                
    except KeyboardInterrupt:
        print("\n🛑 Script interrupted by user")
    except Exception as e:
        print(f"❌ Fatal error: {type(e).__name__}")
        print(f"Message: {str(e)[:100]}")
        traceback.print_exc()
    finally:
        # Clean up Playwright resources
        if playwright:
            playwright.stop()
        print("✅ Playwright resources cleaned up")


# --- Execução ---
if __name__ == "__main__":
    start_time = time.perf_counter()
    executar_scraping()
    elapsed = time.perf_counter() - start_time
    print(f"⏱️ Tempo total de execução: {elapsed:.2f} segundos")