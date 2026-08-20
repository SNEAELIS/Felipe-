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


# --- Extrai os processos das propostas ---
def access_proposal_details(page: str, numero_processo: str): #  <--- MOD THIS ONE
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

# --- Extrai os dados de cada processo ---
def extract_received_processos(page) -> pd.DataFrame: #  <--- MOD THIS ONE
    """Extract all process numbers from the received processes table.
    
    Filters out red-colored text, keeping all other colors.
    """
    try:
        # Get all process numbers and filter out red ones in one go
        result = page.evaluate(r'''
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

        return df, skipped_as_dicts

    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"❌ Error extracting received process numbers: {str(e)[:100]}")
        print(f"Error type: {exc_type.__name__} at line {exc_tb.tb_lineno}")
        return pd.DataFrame(columns=['Processo'])


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



# --- Função principal ---
def executar_scraping(): #  <--- MOD THIS ONE
    """Main execution function"""
    
    other_door = '9232'
    
    page = None
    playwright = None

    try:
        if isinstance(other_door, str) and other_door.isdigit():
            page, playwright, browser = conectar_navegador_existente(int(other_door))
        
        if not page:
            print("❌ Failed to connect to browser. Exiting.")
            return
        

        while True:
            try:
                # Main loop
                state = extract_received_processos(page=page)

                if state:
                    for _ in state: # <-- Itera os processos

                    page.click('#lnkControleProcessos')

                
                time.sleep(15)  
               
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