import shutil
import re
import sys
import time
import os
import traceback
import json

from dataclasses import dataclass, fields

from pathlib import Path
from datetime import datetime

import pandas as pd

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError


os.system('cls' if os.name == 'nt' else 'clear')


@dataclass
class WorkbookState:

    def __init__(self):

        self.bloco = None

        self.processos = None

        self.andamento = None

        self.controle = None


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


# --- Conectar ao navegador existente ---
def conectar_navegador_existente(porta: int):
    """Connect to existing Chrome instance via Playwright"""
    try:
        print(f"Tentando conectar ao navegador na porta {porta}...")
        
        playwright = sync_playwright().start()
        
        # Connect to existing Chrome instance
        browser = playwright.chromium.connect_over_cdp(f"http://127.0.0.1:{porta}")
        
        # Get the first page or create a new one
        if browser.contexts:
            context = browser.contexts[0]
            if context.pages:
                page = context.pages[0]
            else:
                page = context.new_page()
        else:
            context = browser.new_context()
            page = context.new_page()
        
        # Check if we're on SEI
        switch_to_sei(page)

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
                time.sleep(0.3)
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
                print(f"Total de {len(resultados)} linhas extraidas do BLOCO DE ASSINATURA")
                return pd.DataFrame(resultados)
            else:
                print("Nenhuma linha encontrada para salvar")
                        
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred: {str(e)[:100]}")
            print(f"Error type: {exc_type.__name__}")
            print(f"Line number: {exc_tb.tb_lineno}")
            
    except Exception as e:
        exc_tb = sys.exc_info()
        print(f"❌Error occurred at line: {exc_tb.tb_lineno}")
        print(f"{type(e).__name__} - {str(e)[:150]}")
        return None


# --- Extrai os processos das propostas ---
def extrair_dados_propostas(page, arquivo_excel: str):
    """Extract data from proposals"""
    try:
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
                    continue

                linha = linhas_bloco[idx]
                cells = linha.query_selector_all('td')
                if len(cells) < 2:
                    continue

                numero = cells[1].text_content().strip() if cells[1] else ''
                link = cells[1].query_selector('a')
                if not link:
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
                        'Seq.': '',
                        'Processo': '',
                        'Documento': '',
                        'Tipo': '',
                        'Assinaturas': '',
                        'Anotações': '',
                        'Ações': '',
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
            print(f"Total de {len(resultados)} linhas encontradas nas PROPOSTAS")
            return pd.DataFrame(resultados)

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
        start_time = time.time()
        frames_loaded = False
        last_frame_count = 0
        
        while time.time() - start_time < timeout / 1000:
            try:
                frames = page.frames
                current_frame_count = len(frames)
                
                if current_frame_count != last_frame_count:
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
                        frames_loaded = True
                        break
                else:
                    # If no expected frames, wait for at least the iframes in DOM
                    try:
                        iframes = page.query_selector_all('iframe')
                        if len(iframes) > 0 and len(frames) > 1:
                            frames_loaded = True
                            break
                    except Exception as e:
                        # If query_selector_all fails, continue
                        pass
                
                time.sleep(1)
                
            except Exception as e:
                exc_type, exc_value, exc_tb = sys.exc_info()
                print(f"⚠️ Error checking frames: {str(e)[:50]}")
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
                    page.evaluate(r"""
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
    deadline = time.time() + timeout / 1000.0
    last_error = None

    while time.time() < deadline:
        try:
            # Direct frame by name or id
            frame = page.frame(name='ifrVisualizacao')
            if frame and not frame.is_detached():
                return frame

            iframe_element = page.query_selector('iframe[name="ifrVisualizacao"], iframe#ifrVisualizacao')
            if iframe_element:
                frame = iframe_element.content_frame()
                if frame and not frame.is_detached():
                    return frame

            parent_iframe = page.query_selector('iframe[name="ifrConteudoVisualizacao"], iframe#ifrConteudoVisualizacao')
            if parent_iframe:
                parent_frame = parent_iframe.content_frame()
                if parent_frame and not parent_frame.is_detached():
                    child_iframe = parent_frame.query_selector('iframe[name="ifrVisualizacao"], iframe#ifrVisualizacao')
                    if child_iframe:
                        frame = child_iframe.content_frame()
                        if frame and not frame.is_detached():
                            return frame

            for f in page.frames:
                if f.is_detached():
                    continue
                if f.name == 'ifrVisualizacao' or 'visualizacao' in str(f.name).lower() or 'arvore_processar_html' in f.url:
                    return f

        except Exception as e:
            last_error = e

        time.sleep(0.5)

    if last_error:
        print(f"⚠️ wait_for_ifr_visualizacao_frame timeout. Last error: {type(last_error).__name__}: {str(last_error)[:50]}")
    else:
        print("⚠️ wait_for_ifr_visualizacao_frame timeout without finding the frame.")
    return None


# --- Extrai os dados da tabela de histórico ---
def extract_history_from_nested_frame(page, numero_processo) -> list[dict]:
    """Extract history table from the nested ifrVisualizacao frame"""
    
    try:
        viz_frame = wait_for_ifr_visualizacao_frame(page, timeout=20000)
        if not viz_frame:
            print("❌ ifrVisualizacao frame not found")
            print("\nAll available frames:")
            return None

        # Wait for content inside ifrVisualizacao
        try:
            viz_frame.wait_for_selector('table, .infraAreaTelad, #tblHistorico, div.infraAreaTelad, tbody', timeout=15000)
            if viz_frame.is_detached():
                viz_frame = wait_for_ifr_visualizacao_frame(page, timeout=10000)
                if not viz_frame:
                    return None
        except Exception as e:
            viz_frame = wait_for_ifr_visualizacao_frame(page, timeout=10000)
            if not viz_frame:
                return None

        # Step 6: Extract the table data
        try:
            js_source = r"""
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
            js_source = r"""
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

        #print(f"✅ DEBUGG !!! Extracted {len(results)} history entries for process {numero_processo}  DEBUGG !!!")
            
        return pd.DataFrame(results)
        
    except Exception as e:
        return None


# --- Extrai detalhes das propostas (andamentos) ---
def get_proposal_details(page, arquivo_excel: str, sheet_name: str, processos=None ):
    """Extract proposal tracking details and save to specified sheet_name"""
    
    columns = ['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição']
    
    if sheet_name == 'Andamento':
        df = processos
        pattern = r'^\d{5}\.\d{6}/\d{4}-\d{2}$'
        df = df[df['Processo'].str.match(pattern, na=False)]
        df = df.drop_duplicates(subset=['Processo'])
        numeros_processo = df['Processo'].tolist()
        print(f"🔎 Found {len(numeros_processo)} unique process numbers in 'Andamento' sheet")
    else:
        page.click('#lnkControleProcessos')
        source, new_processes = extract_received_processos(page=page)
        numeros_processo = source['Processo'].tolist()
        print(source)
        print(f"🔎 Found {len(numeros_processo)} process numbers in the source sheet")
        
   
    all_results_df = pd.DataFrame(columns=columns)
    for numero_processo in numeros_processo:
        try:
            wait_for_frames_to_load(page=page, timeout=30000)

            # Go to default content (top frame)
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
            wait_for_frames_to_load(page=page, expected_frame_names=None, timeout=30000) 
            
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
                    if len(resultados) > 0:
                        df_resultados = resultados.copy()
                        # Ensure columns match
                        for col in columns:
                            if col not in df_resultados.columns:
                                df_resultados[col] = ''
                        
                        df_resultados = df_resultados[columns]
                        all_results_df = pd.concat([all_results_df, df_resultados], ignore_index=True)

                        #print(f"✅ {len(resultados)} registros de andamento coletados para {numero_processo}")
                        break
                    else:
                        time.sleep(1)
                        attempt += 1
                
            except Exception as e:
                continue
                
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error processing {numero_processo}: {str(e)[:100]}")
            print(f"Error type: {exc_type.__name__}")
            print(f"Line number: {exc_tb.tb_lineno}")
            sys.exit()

    if not  sheet_name == 'Andamento':
        if new_processes:
            df_new = pd.DataFrame(new_processes)
            # Ensure columns match
            for col in columns:
                if col not in df_new.columns:
                    df_new[col] = ''
            
            # Keep only the needed columns
            df_new = df_new[columns]
            # Concatenate to main DataFrame
            all_results_df = pd.concat([all_results_df, df_new], ignore_index=True)
            print(f"✅ Adicionados {len(new_processes)} processos não visualizados (vermelhos)")

    if not all_results_df.empty:
        return all_results_df
    else:
        print(f"⚠️ Nenhum registro coletado. Nenhuma gravação realizada.")

# --- Extrai os dados de cada processo ---
def extract_received_processos(page) -> pd.DataFrame:
    """Extract all process numbers from the received processes table.
    
    Filters out red-colored text, keeping all other colors.
    """
    try:
        # Get all process numbers and filter out assigned ones in one go
        result = page.evaluate(r"""
        () => {
        const rows = document.querySelectorAll('#tblProcessosRecebidos tbody tr');
        const processosNovos = [];
        const skipped = [];
        
        for (let i = 1; i < rows.length; i++) {
            const cells = rows[i].querySelectorAll('td');
            
            if (cells.length >= 3) {
                const cellProcesso = cells[2];  // Column with the process number link
                const link = cellProcesso.querySelector('a');
                
                if (!link) continue;
                
                // Get assignment text from the LAST column
                const lastCell = cells[cells.length - 1];
                const rawAtribuicao = lastCell ? lastCell.textContent.trim() : '';
                
                // Clean up parentheses and whitespace:
                const atribuidoA = rawAtribuicao.replace(/[()]/g, '').trim();
                
                const text = link.textContent.trim().replace(/\s+/g, '');
                
                if (text && text.includes('.')) {
                    if (atribuidoA === '') {
                        // Unassigned process - include it
                        processosNovos.push(text);
                    } else {
                        // Assigned process - skip it
                        skipped.push(text);
                    }
                }
            }
        }
        
        return {
            processos: processosNovos,
            skipped: skipped
        };
        }
        """)

        skipped_as_dicts = []
        print(f"🔎 Found {len(result['skipped'])} assigned processes (skipped)")
        print(f"🔎 Found {len(result['processos'])} unassigned processes (included)")

        for process in result['processos']:
            skipped_as_dicts.append({
                'Processo': process,
                'Data/Hora': 'novo',
                'Unidade': 'novo',
                'Usuário': 'novo',
                'Descrição': 'novo'
                # Add any other fields that your history extraction returns
            })
        
        df = pd.DataFrame({'Processo': result['processos']})

        return df, skipped_as_dicts

    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"❌ Error extracting received process numbers: {str(e)[:100]}")
        print(f"Error type: {exc_type.__name__} at line {exc_tb.tb_lineno}")
        return pd.DataFrame(columns=['Processo']), []


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
                return frame
        
        # Try by URL pattern
        for frame in page.frames:
            if 'conteudo' in frame.url or 'procedimento' in frame.url:
                return frame
        
        print(f"⚠️ Could not find frame: {iframe_id}")
        return None
        
    except Exception as e:
        print(f"Error finding iframe: {type(e).__name__} - {str(e)[:80]}")
        return None

# --- Salva os dados no excel ---
def salvar_excel(arquivo: str | Path, dados) -> bool:
    """
    Save data from dataclass to Excel file, overwriting all sheets
    
    Args:
        arquivo: Path to Excel file
        dados: WorkbookState instance with bloco, processos, andamento, controle attributes
    
    Returns:
        bool: True if successful, False otherwise
    """

    SHEET_CONFIG = {
    'Bloco': {
        'columns': ['Número', 'Sinalizações', 'Atribuição', 'Estado', 
                   'Geradora', 'Disponibilização', 'Grupo', 'Descrição', 'Ações']
    },
    'Processos': {
        'columns': ['Número', 'Seq.', 'Processo', 'Documento', 'Tipo', 
                   'Assinaturas', 'Anotações', 'Ações']
    },
    'Andamento': {
        'columns': ['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição']
    },
    'Controle de Processo': {
        'columns': ['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição']
    }
    }

    total_rows = 0

    try:
        arquivo = Path(arquivo)
        
        # Get data from dataclass (works with __init__ attributes)
        sheet_mapping = {
            'Bloco': getattr(dados, 'bloco', None),
            'Processos': getattr(dados, 'processos', None),
            'Andamento': getattr(dados, 'andamento', None),
            'Controle de Processo': getattr(dados, 'controle', None)
        }
        
        # Debug: Print what we received
        print("\n📊 Data received in salvar_excel:")
        for sheet_name, df in sheet_mapping.items():
            if df is not None and isinstance(df, pd.DataFrame) and not df.empty:
                print(f"  ✅ {sheet_name}: {len(df)} rows, {len(df.columns)} columns")
            elif df is not None and isinstance(df, pd.DataFrame) and df.empty:
                print(f"  ⚠️  {sheet_name}: Empty DataFrame")
            else:
                print(f"  ❌ {sheet_name}: {type(df).__name__ if df is not None else 'None'}")

        # Prepare final data for all sheets
        final_data = {}
        
        for sheet_name, df in sheet_mapping.items():
            # Get expected columns for this sheet
            expected_cols = SHEET_CONFIG[sheet_name]['columns']
            
            # Check if df is a valid DataFrame
            if df is not None and isinstance(df, pd.DataFrame):
                # Copy to avoid modifying original
                df_clean = df.copy()
                
                # Add missing columns with empty strings
                for col in expected_cols:
                    if col not in df_clean.columns:
                        df_clean[col] = ''
                
                # Keep only expected columns in correct order
                # Only keep columns that exist in the DataFrame
                existing_cols = [col for col in expected_cols if col in df_clean.columns]
                if existing_cols:
                    df_clean = df_clean[existing_cols]
                else:
                    df_clean = pd.DataFrame(columns=expected_cols)
            else:
                # Create empty dataframe with correct columns
                df_clean = pd.DataFrame(columns=expected_cols)
            
            final_data[sheet_name] = df_clean
            total_rows += len(df_clean)
            
            print(f"✅ {len(df_clean)} linhas preparadas para '{sheet_name}'")
        
        # Ensure all sheets exist even if not provided
        for sheet_name in SHEET_CONFIG.keys():
            if sheet_name not in final_data:
                final_data[sheet_name] = pd.DataFrame(columns=SHEET_CONFIG[sheet_name]['columns'])
                print(f"ℹ️  Criada planilha vazia: '{sheet_name}'")
        
        # Create directory if it doesn't exist
        arquivo.parent.mkdir(parents=True, exist_ok=True)
        
        # Save all sheets to file (overwrites existing file)
        with pd.ExcelWriter(arquivo, engine='openpyxl', mode='w') as writer:
            for sheet_name, df in final_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Success summary
        print(f"\n✅ Dados salvos com sucesso em {arquivo}")
        print(f"   Total: {total_rows} linhas em {len(final_data)} planilhas")
        
        return True
        
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"❌ Erro ao salvar no Excel: {type(e).__name__}")
        print(f"   Linha: {exc_tb.tb_lineno}")
        print(f"   Mensagem: {str(e)[:100]}")
        import traceback
        traceback.print_exc()
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

# --- Salva dados concatenados em Excel com data para avaliação de performance diária ---
def salvar_excel_com_data(arquivo: str | Path, dados) -> bool:
    """
    Save data from dataclass to Excel file with daily timestamp in filename.
    If file for today exists, it concatenates new data with existing data.
    
    Args:
        arquivo: Base path for the Excel file (will have date added)
        dados: WorkbookState instance with bloco, processos, andamento, controle attributes
    
    Returns:
        bool: True if successful, False otherwise
    """
    
    def _get_id_columns(sheet_name: str) -> list:
        """
        Get the columns that uniquely identify a row for each sheet.
        Used for deduplication when concatenating data.
        """
        ID_COLUMNS = {
            'Bloco': ['Número'],  # Process number should be unique in block
            'Processos': ['Processo', 'Seq.'],  # Process + sequence number
            'Andamento': ['Processo', 'Data/Hora', 'Descrição'],  # Unique entry in history
            'Controle de Processo': ['Processo', 'Data/Hora', 'Descrição']  # Unique entry in history
        }
        
        return ID_COLUMNS.get(sheet_name, [])

    SHEET_CONFIG = {
        'Bloco': {
            'columns': ['Número', 'Sinalizações', 'Atribuição', 'Estado', 
                       'Geradora', 'Disponibilização', 'Grupo', 'Descrição', 'Ações']
        },
        'Processos': {
            'columns': ['Número', 'Seq.', 'Processo', 'Documento', 'Tipo', 
                       'Assinaturas', 'Anotações', 'Ações']
        },
        'Andamento': {
            'columns': ['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição']
        },
        'Controle de Processo': {
            'columns': ['Processo', 'Data/Hora', 'Unidade', 'Usuário', 'Descrição']
        }
    }
    
    total_rows = 0
    new_rows = 0
    existing_rows = 0
    
    try:
        # Convert to Path object
        arquivo = Path(arquivo)
        
        # Generate date stamp (YYYYMMDD)
        date_stamp = datetime.now().strftime("%Y%m%d")
        
        # Create the filename with date
        base_name = arquivo.stem
        extension = arquivo.suffix
        daily_filename = arquivo.parent / f"{base_name}_{date_stamp}{extension}"
        
        #print(f"📁 Daily file: {daily_filename}")
        
        # Get current data from dataclass
        sheet_mapping = {
            'Bloco': getattr(dados, 'bloco', None),
            'Processos': getattr(dados, 'processos', None),
            'Andamento': getattr(dados, 'andamento', None),
            'Controle de Processo': getattr(dados, 'controle', None)
        }
        
        # Check if daily file exists
        if daily_filename.exists():
            #print(f"📂 Daily file exists. Loading existing data...")
            
            # Load existing data from all sheets
            existing_data = {}
            try:
                with pd.ExcelFile(daily_filename) as xls:
                    for sheet_name in SHEET_CONFIG.keys():
                        if sheet_name in xls.sheet_names:
                            df_existing = pd.read_excel(xls, sheet_name=sheet_name)
                            existing_data[sheet_name] = df_existing
                            existing_rows += len(df_existing)
                            #print(f"  📊 {sheet_name}: {len(df_existing)} existing rows")
                        else:
                            existing_data[sheet_name] = pd.DataFrame(columns=SHEET_CONFIG[sheet_name]['columns'])
            except Exception as e:
                print(f"⚠️ Could not read existing file: {e}")
                existing_data = {}
        else:
            #print(f"📄 Daily file does not exist. Creating new file...")
            existing_data = {}
        
        # Prepare final data for all sheets
        final_data = {}
        
        for sheet_name, df in sheet_mapping.items():
            # Get expected columns for this sheet
            expected_cols = SHEET_CONFIG[sheet_name]['columns']
            
            # Clean the new data
            if df is not None and isinstance(df, pd.DataFrame):
                df_clean = df.copy()
                
                # Add missing columns with empty strings
                for col in expected_cols:
                    if col not in df_clean.columns:
                        df_clean[col] = ''
                
                # Keep only expected columns
                existing_cols = [col for col in expected_cols if col in df_clean.columns]
                if existing_cols:
                    df_clean = df_clean[existing_cols]
                else:
                    df_clean = pd.DataFrame(columns=expected_cols)
            else:
                df_clean = pd.DataFrame(columns=expected_cols)
            
            # Get existing data for this sheet
            df_existing = existing_data.get(sheet_name, pd.DataFrame(columns=expected_cols))
            
            # Ensure existing data has all columns
            for col in expected_cols:
                if col not in df_existing.columns:
                    df_existing[col] = ''
            
            # Keep only expected columns in existing data
            df_existing = df_existing[expected_cols] if not df_existing.empty else pd.DataFrame(columns=expected_cols)
            
            # Check for duplicates and concatenate
            if not df_clean.empty:
                # Determine which columns to use for deduplication
                id_columns = _get_id_columns(sheet_name)
                
                if not df_existing.empty:
                    # Combine existing and new data
                    combined = pd.concat([df_existing, df_clean], ignore_index=True)
                    
                    # Remove duplicates based on ID columns
                    if id_columns:
                        # Keep first occurrence (existing data)
                        combined = combined.drop_duplicates(subset=id_columns, keep='first')
                        new_rows_added = len(df_clean) - (len(combined) - len(df_existing))
                    else:
                        # If no ID columns, just concatenate all
                        combined = df_clean
                        new_rows_added = len(df_clean)
                    
                    final_data[sheet_name] = combined
                    new_rows += new_rows_added
                    total_rows += len(combined)
                    
                else:
                    # No existing data, just use new data
                    final_data[sheet_name] = df_clean
                    new_rows += len(df_clean)
                    total_rows += len(df_clean)
            else:
                # No new data, keep existing
                final_data[sheet_name] = df_existing
                total_rows += len(df_existing)
        
        # Ensure all sheets exist
        for sheet_name in SHEET_CONFIG.keys():
            if sheet_name not in final_data:
                final_data[sheet_name] = pd.DataFrame(columns=SHEET_CONFIG[sheet_name]['columns'])
        
        # Create directory if it doesn't exist
        daily_filename.parent.mkdir(parents=True, exist_ok=True)
        
        # Save all sheets to file
        with pd.ExcelWriter(daily_filename, engine='openpyxl', mode='w') as writer:
            for sheet_name, df in final_data.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
        return True
        
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"❌ Erro ao salvar no Excel com data: {type(e).__name__}")
        print(f"   Linha: {exc_tb.tb_lineno}")
        print(f"   Mensagem: {str(e)[:100]}")
        import traceback
        traceback.print_exc()
        return False

# --- Função principal ---
def executar_scraping():
    """Main execution function"""
    
    arquivo_fonte = r"C:\Users\felipe.rsouza\Automação SNEAELIS\Dashboard sei DB\DB_sei_se.xlsx"
    arquivo_destino = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\Controle_SEI\DB_sei_se.xlsx"
    arquivo_destino_2 = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Mateus - SEI\DB_sei_se.xlsx"
    other_door = '9230'
    
    page = None
    playwright = None
    state = WorkbookState()

    try:
        if isinstance(other_door, str) and other_door.isdigit():
            page, playwright, browser = conectar_navegador_existente(int(other_door))
        
        if not page:
            print("❌ Failed to connect to browser. Exiting.")
            return
        

        while True:
            try:
                # Check current time to decide if we should save
                now = datetime.now()
                current_hour = now.hour
                
                # Only save between 7:00 and 21:00
                should_save = 7 <= current_hour < 21

                # Main loop
                acessa_bloco_ass(page=page)

                state.bloco = extrair_dados_bloco(page=page,
                                     arquivo_excel=arquivo_fonte
                                     )
                
                state.processos = extrair_dados_propostas(page=page,
                                        arquivo_excel=arquivo_fonte
                                        )
 
                state.andamento = get_proposal_details(page=page,
                                     arquivo_excel=arquivo_fonte,
                                     sheet_name='Andamento',
                                     processos=state.processos
                                     )
                
                state.controle = get_proposal_details(page=page,
                                     arquivo_excel=arquivo_fonte,
                                     sheet_name='Controle de Processo'
                                     )

                salvar_excel(arquivo=arquivo_fonte, dados=state)

                if should_save:
                    salvar_excel_com_data(arquivo=arquivo_fonte, dados=state)

                page.click('#lnkControleProcessos')

                shutil.copy(arquivo_fonte, arquivo_destino)
                shutil.copy(arquivo_fonte, arquivo_destino_2)
                print(f"\n✅ Copied file to {arquivo_destino}\n")
                
                time.sleep(5)  
               
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