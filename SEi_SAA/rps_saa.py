from itertools import count
import re
import sys
import time
import os
import traceback

from playwright.sync_api import sync_playwright, Page, expect, TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError

from typing import Optional, Dict

from streamlit import iframe

os.system('cls' if os.name == 'nt' else 'clear')

# ─── CONSTANTES ───────────────────────────────────────────────────────────────

BLOCO_ASSINATURA_RP = '435153'#'432067'  # ANDREI RODRIGUES - RP e GRU
BLOCO_INTERNO_RP = '434283'     # RP e GRU pelo Robôzinho


# ─── FUNÇÕES DE CONEXÃO ──────────────────────────────────────────────────────

def switch_to_sei(page):
    """Switch to the correct SEI page/tab"""
    target_url = "sei.mds.gov.br/sei/controlador.php?acao=procedimento"
    #target_url = "sei.mds.gov.br/sei/controlador.php?acao=editor_montar"

    
       
    context = page.context
    all_pages = context.pages
    print(f"🔄 Switching to SEI tab. Total open tabs: {len(all_pages)}")
    
    for p in all_pages:
        if target_url in p.url:
            print(f"🎯 Found and switched to: {p.url}")
            return p
    
    print("❌ Target URL not found in any open tabs.")
    return None


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
        
        page = switch_to_sei(page)
        print("Current URL:", page.url)
        return page, playwright, browser
        
    except Exception as e:
        msg = "Erro ao conectar. Verifique se o Chrome está aberto com depuração."
        print("❌", msg)
        print(f'{type(e).__name__}\n\n{str(e)[:100]}')
        return None, None, None


# ─── UTILITÁRIOS ─────────────────────────────────────────────────────────────

def formato_padrao(num_sei: str):
    """Convert process number to standard format: 12345.678901/2024-00"""
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


def wait_for_frames_to_load(page, expected_frame_names=None, timeout=30000):
    """Wait for frames to be loaded and available in page.frames"""
    try:        
        start_time = time.time()
        frames_loaded = False
        
        """ while time.time() - start_time < timeout / 500:
            try:
                frames = page.frames
                print(f"✅ Frame '{name}' found.")

                if expected_frame_names:
                    found_frames = []
                    for name in expected_frame_names:
                        try:
                            frame = page.frame(name=name)
                            if frame:
                                found_frames.append(name)
                        except Exception:
                            continue
                    
                    if len(found_frames) == len(expected_frame_names):
                        frames_loaded = True
                        break
                else:
                    try:
                        iframes = page.query_selector_all('iframe')
                        if len(iframes) > 0 and len(frames) > 1:
                            frames_loaded = True
                            break
                    except Exception:
                        pass
                
                time.sleep(0.5)
                
            except Exception:
                time.sleep(0.5)
                print(f"⏳Attempting to wait for frames to load... ({int(time.time() - start_time)}s elapsed)")

                continue """
        
        if not frames_loaded:
            # Try to force load
            try:
                page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
                time.sleep(0.3)
                page.evaluate("window.scrollTo(0, 0);")
            except Exception:
                pass
        
        try:
            return page.frames
        except:
            return []
    
    except Exception as e:
        print(f"❌ wait_for_frames_to_load failed: {str(e)[:100]}")
        return []


def debug_all_frames(page, single_frame: str = ''):
    """Debug all frames in Playwright"""
    print("\n" + "="*60)
    print("DEBUGGING ALL FRAMES IN PLAYWRIGHT")
    print("="*60)
    
    frames = page.frames
    print(f"\n📊 Total frames in page.frames: {len(frames)}")
    
    for idx, frame in enumerate(frames):
        if single_frame and frame.name != single_frame:
            continue
        print(f"\n--- FRAME {idx} ---")
        print(f"  Name: {frame.name}")
        print(f"  URL: {frame.url[:80] if frame.url else 'None'}")
        
        try:
            title = frame.evaluate('document.title')
            print(f"  Title: {title}")
        except Exception as e:
            print(f"  ❌ Cannot access frame: {e}")
    
    iframe_elements = page.query_selector_all('iframe')
    print(f"\n📊 Iframe elements found: {len(iframe_elements)}")
    
    return frames, iframe_elements


def debug_page_info(page, label: str = ''):
    """
    Debug helper: dump everything about the current page (URL, title, open tabs,
    rendered text, frame structure, key SEI element presence).

    Call this manually anywhere, e.g.:
        debug_page_info(page, "após abrir o popup do editor")

    Never raises — each inspection is individually guarded.
    """
    print("\n" + "="*70)
    print(f"[DEBUG] {label or 'PÁGINA ATUAL'}")
    print("="*70)

    try:
        print(f"\n🌐 URL: {page.url}")
    except Exception as e:
        print(f"\n🌐 URL: ❌ {type(e).__name__}: {str(e)[:100]}")

    try:
        print(f"📄 Título: {page.title()}")
    except Exception as e:
        print(f"📄 Título: ❌ {type(e).__name__}: {str(e)[:100]}")

    try:
        print(f"🔄 readyState: {page.evaluate('document.readyState')}")
    except Exception as e:
        print(f"🔄 readyState: ❌ {type(e).__name__}: {str(e)[:100]}")

    # ── Abas abertas ──────────────────────────────────────────────────────────
    try:
        all_pages = page.context.pages
        try:
            current_idx = all_pages.index(page)
        except ValueError:
            current_idx = -1
        print(f"\n📑 Abas abertas: {len(all_pages)} (esta é a aba #{current_idx + 1 if current_idx >= 0 else '?'})")
        for i, p in enumerate(all_pages):
            marker = "  ◀── ESTA" if p is page else ""
            try:
                print(f"   [{i + 1}] {p.url}")
            except Exception as e:
                print(f"   [{i + 1}] ❌ {type(e).__name__}: {str(e)[:80]}")
            try:
                print(f"       título: {p.title()}{marker}")
            except Exception as e:
                print(f"       título: ❌ {type(e).__name__}: {str(e)[:80]}{marker}")
    except Exception as e:
        print(f"\n📑 Abas abertas: ❌ {type(e).__name__}: {str(e)[:100]}")

    # ── Conteúdo renderizado ──────────────────────────────────────────────────
    try:
        body_text = page.evaluate('document.body ? document.body.innerText : ""') or ''
        body_text = body_text.strip().replace('\n', ' | ')
        print(f"\n📝 Conteúdo da página ({len(body_text)} caracteres):")
        print(f"   {body_text[:500]}{'...' if len(body_text) > 500 else ''}")
    except Exception as e:
        print(f"\n📝 Conteúdo da página: ❌ {type(e).__name__}: {str(e)[:100]}")

    # ── Frames / iframes ──────────────────────────────────────────────────────
    try:
        frames = page.frames
        print(f"\n🧩 Frames ({len(frames)}):")
        for i, frame in enumerate(frames):
            try:
                title = frame.evaluate('document.title')
            except Exception:
                title = '?'
            print(f"   [{i}] name='{frame.name}' url='{frame.url[:120]}' title='{title[:60]}'")
    except Exception as e:
        print(f"\n🧩 Frames: ❌ {type(e).__name__}: {str(e)[:100]}")

    # ── Indicadores de estrutura (elementos-chave do SEI) ────────────────────
    try:
        key_selectors = {
            '#lnkControleProcessos': 'link Controle de Processos',
            '#txtPesquisaRapida': 'campo pesquisa rápida',
            '#ifrArvore': 'iframe ifrArvore',
            '#ifrVisualizacao': 'iframe ifrVisualizacao',
            '#ifrConteudoVisualizacao': 'iframe ifrConteudoVisualizacao',
            '.ck-editor__editable': 'editor CKEditor',
            '#txaEditor_935': 'textarea txaEditor_935',
            '#selBloco': 'select selBloco',
            '#sbmIncluir': 'botão sbmIncluir',
        }
        print(f"\n🔍 Indicadores de estrutura:")
        for selector, desc in key_selectors.items():
            try:
                found = len(page.query_selector_all(selector))
            except Exception:
                found = -1
            status = '✅' if found > 0 else ('⚠️' if found == -1 else '❌')
            print(f"   {status} {desc} ({selector}): {found if found >= 0 else 'erro'}")

        try:
            links_count = len(page.query_selector_all('a'))
            tables_count = len(page.query_selector_all('table'))
            print(f"   📊 <a> links: {links_count} | <table> tabelas: {tables_count}")
        except Exception as e:
            print(f"   📊 contagem de elementos: ❌ {type(e).__name__}: {str(e)[:100]}")
    except Exception as e:
        print(f"\n🔍 Indicadores de estrutura: ❌ {type(e).__name__}: {str(e)[:100]}")

    print("="*70)


def detect_and_extract_editor_content(page: Page) -> Dict[str, any]:
    """
    Detects CKEditor presence and extracts all content with labels.
    
    Returns:
        Dict containing:
        - editor_detected: bool
        - editor_type: str
        - roots: Dict with each root's content and metadata
        - summary: Text summary of all content
    """
    
    result = page.evaluate(
        """
        () => {
            // ============================================
            // 1. DETECT CKEDITOR
            // ============================================
            const hasInfraEditor = typeof window.infraEditor !== 'undefined';
            const isReady = hasInfraEditor && window.infraEditor.isReady === true;
            const hasCkeditor = typeof window.CKEDITOR !== 'undefined';
            
            const editorDetected = hasInfraEditor || hasCkeditor;
            
            // Determine editor type
            let editorType = 'unknown';
            if (hasInfraEditor && isReady) {
                editorType = 'infraEditor (CKEditor 5)';
            } else if (hasInfraEditor) {
                editorType = 'infraEditor (loading)';
            } else if (hasCkeditor) {
                editorType = 'CKEditor (legacy)';
            }
            
            // ============================================
            // 2. GET EDITOR CONFIG
            // ============================================
            let config = null;
            try {
                config = window.INFRA_EDITOR_CONFIG || null;
            } catch (e) {
                // Ignore
            }
            
            // ============================================
            // 3. EXTRACT CONTENT FROM EACH ROOT
            // ============================================
            const rootConfigs = config ? config.rootsAttributes : null;
            const editor = window.infraEditor;
            
            // Define all possible editor roots from your HTML
            const rootIds = [
                'txaEditor_926',
                'txaEditor_929', 
                'txaEditor_932',
                'txaEditor_935',
                'txaEditor_941'
            ];
            
            const roots = {};
            let fullText = '';
            let hasContent = false;
            
            for (const rootId of rootIds) {
                // Get root metadata
                let metadata = {
                    label: rootId,
                    isPrincipal: false,
                    isReadOnly: false,
                    exists: false
                };
                
                // Get config metadata if available
                if (rootConfigs && rootConfigs[rootId]) {
                    metadata.label = rootConfigs[rootId].label || rootId;
                    metadata.isPrincipal = rootConfigs[rootId].principal || false;
                    metadata.isReadOnly = rootConfigs[rootId].somenteLeitura || false;
                }
                
                // Try to get content
                let html = '';
                let text = '';
                let error = null;
                
                // Method 1: CKEditor API
                try {
                    if (editor && editor.isReady) {
                        const rootEditor = editor._editors ? editor._editors.get(rootId) : null;
                        if (rootEditor) {
                            html = rootEditor.getData();
                        } else {
                            html = editor.getData({ rootName: rootId });
                        }
                    }
                } catch (e) {
                    error = e.message;
                }
                
                // Method 2: DOM fallback
                if (!html) {
                    try {
                        const el = document.getElementById(rootId);
                        if (el) {
                            html = el.innerHTML;
                            metadata.exists = true;
                        }
                    } catch (e) {
                        error = error || e.message;
                    }
                }
                
                // Extract plain text from HTML
                if (html) {
                    // Simple HTML to text conversion
                    const temp = document.createElement('div');
                    temp.innerHTML = html;
                    text = temp.textContent || '';
                    // Clean up extra whitespace
                    text = text.replace(/\\s+/g, ' ').trim();
                    hasContent = hasContent || text.length > 0;
                }
                
                roots[rootId] = {
                    ...metadata,
                    html: html || '',
                    text: text || '',
                    hasContent: text && text.length > 0,
                    error: error
                };
                
                if (text) {
                    fullText += `[${metadata.label}] ${text}\\n`;
                }
            }
            
            // ============================================
            // 4. FIND EDITOR CONTAINER IN DOM
            // ============================================
            const editorContainer = document.querySelector('.infra-editor, .infra-editor__editor, [class*="ck-editor"]');
            
            // ============================================
            // 5. DETECT IF WE'RE IN A FRAME
            // ============================================
            const inIframe = window !== window.top;
            
            // ============================================
            // 6. RETURN ALL DATA
            // ============================================
            return {
                editor_detected: editorDetected,
                editor_ready: isReady,
                editor_type: editorType,
                in_iframe: inIframe,
                has_content: hasContent,
                editor_container: editorContainer ? {
                    tag: editorContainer.tagName,
                    classes: editorContainer.className,
                    visible: editorContainer.offsetParent !== null
                } : null,
                config: config ? {
                    hasConfig: true,
                    document_id: config.infra?.idDocumento || null,
                    tipo_editor: config.infra?.tipoEditor || null,
                    roots_count: config.rootsAttributes ? Object.keys(config.rootsAttributes).length : 0
                } : null,
                roots: roots,
                summary: {
                    total_roots: rootIds.length,
                    roots_with_content: Object.values(roots).filter(r => r.hasContent).length,
                    full_text: fullText.trim() || 'No text content found',
                    principal_content: roots['txaEditor_935']?.text || 'No principal content'
                }
            };
        }
        """
    )
    
    return result


# ─── FUNÇÃO 1: Extrair processos novos ────────────────────────────────────────
def extract_received_processos(page) -> set:
    """
    Extract all process numbers that are NEW (not viewed / red / classe processoNaoVisualizado).
    
    Returns:
        set[str]: Set of process numbers (already formatted)
    """
    try:
        # Ensure we're on Controle de Processos page
        # Click on Controle de Processos link
        try:
            page.click('#lnkControleProcessos')
            time.sleep(0.3)
            wait_for_frames_to_load(page, timeout=10000)
        except Exception:
            # Already there maybe
            pass
        
        # Evaluate JavaScript to extract non-visualized processes
        result = page.evaluate(r'''
            () => {
                const rows = document.querySelectorAll('#tblProcessosRecebidos tbody tr');
                const processosNovos = [];
                
                for (let i = 1; i < rows.length; i++) {
                    const cells = rows[i].querySelectorAll('td');
                    
                    if (cells.length >= 3) {
                        const cell = cells[2];  // Column with the process number
                        const link = cell.querySelector('a');
                        
                        if (!link) continue;
                        
                        // Check if it's non-visualized (red class)
                        const hasRedClass = link.classList.contains('processoNaoVisualizado');
                        
                        if (hasRedClass) {
                            const text = link.textContent.trim().replace(/\s+/g, '');
                            const ariaLabel = link.getAttribute('aria-label') || '';
                            const onmouseover = link.getAttribute('onmouseover') || '';
                            
                            if (text && text.includes('.')) {
                                processosNovos.push({
                                    numero: text.replace(/\\s+/g, ''),
                                    aria_label: ariaLabel,
                                    onmouseover: onmouseover,
                                    id_procedimento: null
                                });
                            }
                        }
                    }
                }
                
                return processosNovos;
            }
        ''')
        
        # Format and return set of process numbers
        processos_set = set()
        for p in result:
            formatted = formato_padrao(p['numero'])
            if formatted:
                processos_set.add(formatted)
        
        print(f"🔴 Encontrados {len(processos_set)} processos novos (não visualizados)")
        if processos_set:
            for p in sorted(processos_set)[:5]:
                print(f"   - {p}")
            if len(processos_set) > 5:
                print(f"   ... e mais {len(processos_set) - 5}")
        
        return processos_set
    
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"❌ Error extracting received process numbers: {str(e)[:100]}")
        print(f"   Line: {exc_tb.tb_lineno}")
        return set()


# ─── FUNÇÃO 2: Verificar se processo tem RP + SNEAELIS ───────────────────────

def access_proposal_details(page, numero_processo: str) -> bool:
    """
    Access process details and check if it contains RP-6/7/8 AND SNEAELIS.
    
    Flow:
    1. Search for the process number
    2. In ifrArvore, find and click last link with "Formulário Empenho/Descentralização Orçamentária"
    3. In ifrVisualizacao, extract text and check for RP + SNEAELIS
    
    Returns:
        bool: True if process matches RP + SNEAELIS criteria
    """

    def assign_process(page):
        print(f"🔄 Atribuindo processo {numero_processo} a usuário...")
        try:
            frame_arvore = page.frame(name='ifrArvore')
            frame_arvore.locator("a").nth(1).click()
            
            wait_for_frames_to_load(page, expected_frame_names=['ifrConteudoVisualizacao'], timeout=1000)

            iframe_visualizacao = page.frame(name='ifrConteudoVisualizacao')
            btn_incluir = iframe_visualizacao.get_by_title("Atribuir Processo")
            if btn_incluir:
                btn_incluir.click()

                # Select user
                (page.frame_locator('iframe[id*="ConteudoVisualizacao"]')
                    .frame_locator('iframe[id*="Visualizacao"]')
                    .locator('#selAtribuicao')
                    .select_option(label='felipe.rsouza - Felipe Rodrigues de Souza')
                )

                page.frame_locator('iframe[id*="ConteudoVisualizacao"]').frame_locator('iframe[id*="Visualizacao"]').locator("#sbmSalvar").click()
        except Exception as e:
            print(f"⚠️ Erro ao atribuir processo: {str(e)[:80]}")
            sys.exit(1)

        


    print(f"\n🔍 Verificando processo: {numero_processo}")
    
    try:
        # Step 1: Search for the process
        try:
            campo_busca = page.wait_for_selector('#txtPesquisaRapida', timeout=5000)
            campo_busca.fill('')
            campo_busca.fill(numero_processo)
            campo_busca.press('Enter')
        except:
            print(f"⚠️ Campo de busca não encontrado, tentando alternativa...")
            campo_busca = page.wait_for_selector('input[name="txtPesquisaRapida"]', timeout=5000)
            campo_busca.fill('')
            campo_busca.fill(numero_processo)
            campo_busca.press('Enter')


        #start = time.perf_counter()

        wait_for_frames_to_load(page, expected_frame_names=['ifrArvore'], timeout=1000)

        #end = time.perf_counter()
        #print(f"Execution time: {end - start:.4f} seconds")


        # Step 2: Find ifrArvore and click the last link with "Formulário Empenho/Descentralização Orçamentária"

        frame_arvore = page.frame(name='ifrArvore')
        if not frame_arvore:
            time.sleep(0.5)
            frame_arvore = page.frame(url='*arvore*')
        
        if not frame_arvore:
            print(f"⚠️ ifrArvore não encontrado para {numero_processo}")
            return False
        
        try:
            wait_for_frames_to_load(page, expected_frame_names=['ifrArvore'], timeout=1000)

            # Procura por TODOS os <a> dentro de #divArvore
            all_links = frame_arvore.query_selector_all('#divArvore a')
            
            # Palavras-chave para encontrar o link do formulário
            # Sandbox (teste): "Formulário" ou "Geral"
            # Produção: "Formulário Empenho" ou "Descentralização Orçamentária"
            KEYWORDS_SANDBOX = ['Formulário', 'Geral']
            KEYWORDS_PROD = ['Formulário Empenho', 'Descentralização Orçamentária']
            
            matching_links_texts = []
            matching_links_elements = []
            
            for link in all_links:
                try:
                    text = link.text_content().strip()
                    # Tenta match com keywords de produção primeiro, depois sandbox
                    if any(kw in text for kw in KEYWORDS_PROD) or any(kw in text for kw in KEYWORDS_SANDBOX):
                        matching_links_texts.append(text)
                        matching_links_elements.append(link)
                except:
                    continue
            
            if not matching_links_elements:
                print(f"⚠️ Nenhum link de formulário encontrado")
                return False
            
            """ print(f"📄 Encontrados {len(matching_links_elements)} links de formulário:")
            for t in matching_links_texts:
                print(f"   → {t}") """
            
            # Click the LAST matching link
            last_link = matching_links_elements[-1]
            last_link.click()
            print(f"✅ Clicou no último link: '{matching_links_texts[-1][:60]}...'")
            
        except Exception as e:
            print(f"⚠️ Erro ao clicar no link do formulário: {str(e)[:80]}")
            return False
        
        # Step 3: Find ifrVisualizacao and extract text
        try:
            # 1. Wait for the target element inside the nested frames
            #    This replaces both the frame retrieval and the wait_for_selector.
            target_selector = 'xpath=/html/body/div[1]'
            (
                page
                .frame_locator('iframe[id*="ConteudoVisualizacao"]')
                .frame_locator('iframe[id*="Visualizacao"]')
                .locator(target_selector)
                .wait_for(state='attached', timeout=10000)
            )
            print(f"✅ iframe visualização carregado, extraindo texto...")
            
            # 2. Extract text (again using locator chain)
            full_text = (
                page
                .frame_locator('iframe[id*="ConteudoVisualizacao"]')
                .frame_locator('iframe[id*="Visualizacao"]')
                .locator('body')
                .inner_text(timeout=10000)
            )

            #print(f"📝 Texto extraído ({len(full_text)} caracteres)")
            
            # Check for RP-6, RP-7, RP-8 AND SNEAELIS
            has_rp = bool(re.search(r'RP-[678]', full_text))
            has_sneaelis = 'SNEAELIS' in full_text
            has_rp2 = bool(re.search(r'RP-[2]', full_text))  # your second pattern
            has_abcd = 'Pagamento de Oficial de Controle de Dopagem' in full_text  # your second text to search
            is_rp6 = bool(re.search(r'RP-[678]', full_text))

            
            if (has_rp and has_sneaelis) or (has_rp2 and has_abcd) or (is_rp6):
                assign_process(page) 
                time.sleep(0.5)
                return True
            else:
                print(f"❌ Processo {numero_processo} não atende aos critérios")
                if not has_rp:
                    print(f"   → RP-6/7/8 não encontrado")
                if not has_sneaelis:
                    print(f"   → SNEAELIS não encontrado")
                return False
        
        except Exception as e:
            print(f"⚠️ Erro ao extrair texto do ifrVisualizacao: {str(e)[:80]}")
            return False
    
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"❌ Erro em access_proposal_details: {str(e)[:100]}")
        print(f"   Line: {exc_tb.tb_lineno}")
        return False


# ─── FUNÇÃO 3: Criar despacho ────────────────────────────────────────────────

def criar_despacho(page, context, numero_processo: str) -> bool:
    """
    Create a despacho (dispatch document) inside the process.
    
    The CKEditor opens in a new window (popup). We need to:
    1. Click "Incluir Documento" in ifrConteudoVisualizacao
    2. Select "Despacho" from the series table
    3. Wait for new window with CKEditor
    4. Edit the text in txaEditor_935 (Corpo do Texto)
    5. Click "Público" radio button
    6. Save and close
    
    Returns:
        bool: True if successful
    """

    print(f"\n📝 Criando despacho para processo: {numero_processo}") 

    try:
        # Step 1: Switch to ifrConteudoVisualizacao
        frame_conteudo = page.frame(name='ifrConteudoVisualizacao')
        

        if not frame_conteudo:
            frame_conteudo = page.frame(url='*arvore_visualizar*')
        
        if not frame_conteudo:
            print("⚠️ ifrConteudoVisualizacao não encontrado")
            return False
        
        # Wait for the tree to load
        frame_conteudo.wait_for_selector('#divArvoreAcoes', timeout=10000)
               
        # Step 2: Click "Incluir Documento" (first link in divArvoreAcoes)
        try:
            incluir_btn = frame_conteudo.locator('xpath=//*[@id="divArvoreAcoes"]/a[1]')
            time.sleep(2)
            incluir_btn.evaluate("el => el.click()")
               
        except Exception as e:
            print(f"⚠️ Erro ao clicar em Incluir Documento: {str(e)[:100]}")
            return False
        
        # Step 3: Find "Despacho" in the series table
        try:
            # Find the row with "Despacho" text and click it
            despacho_link = (page
                            .frame_locator('iframe[id*="ConteudoVisualizacao"]')
                            .frame_locator('iframe[id*="Visualizacao"]')
                            .locator('#tblSeries tr', has_text='Despacho')
                            .locator('a')
                            .first
            )
            despacho_link.evaluate('() => window.escolher(5)')
            
            time.sleep(0.3)  # Short wait for new window
            
        except Exception as e:
            print(f"⚠️ Erro ao selecionar Despacho: {str(e)[:80]}")
            return False

        try:
            frame_chain = (page
                            .frame_locator('iframe[id*="ConteudoVisualizacao"]')
                            .frame_locator('iframe[id*="Visualizacao"]')
            )

            # Click the "Público" label
            frame_chain.locator('label:has-text("Público")').click()

            # Click the "Salvar" button
            frame_chain.locator('#divInfraBarraComandosInferior #btnSalvar').click()
        except Exception as e:
            print(f"⚠️ Erro ao salvar: {str(e)}")

        # Step 4: Wait for new window/popup with CKEditor
        new_page = None
        try:
            # Listen for new page
            with context.expect_page(timeout=15000) as page_catcher:
                # The new window should already be opening, but we wait
                pass
            
            new_page = page_catcher.value
            
            # Wait for the page to load
            new_page.wait_for_load_state('networkidle', timeout=20000)
            print(f"✅ Nova janela do editor aberta: {new_page.url[:60]}")
            
        except Exception:
            # Check if there's a new page in context
            all_pages = context.pages
            for p in all_pages:
                if 'editor' in p.url.lower() or 'documento_visualizar' in p.url:
                    new_page = p
                    print(f"✅ Editor encontrado em aba existente")
                    break
            
        if not new_page:
            print("⚠️ Não foi possível detectar a nova janela do editor")
            return False
    
        # Step 5: Edit the CKEditor content
        #print(detect_and_extract_editor_content(new_page))  # For debugging purposes
        try:
            # Text for the header
            #new_page = page
            text_to_add = 'Destinatário: MESP/SE/SAA/CGOFC - Coordenação-Geral de Orçamento, Finanças e Contabilidade'
            locator = new_page.locator('#txaEditor_932')
            expect(locator).to_be_editable()

            # Get current text, remove lines containing "Interessado", and append new line
            current_text = locator.inner_text()
            lines = current_text.splitlines()

            keywords = ('Interessado', 'Destinatário')
            lines = [line for line in lines if not any(keyword in line for keyword in keywords)]
            lines.append(text_to_add)
            new_text = '\n'.join(lines)

            locator.fill(new_text)
            locator.blur()  # Trigger change event

            # Text for the body
            # 1. Target the main editor container by ID or class combo
            editor = new_page.locator("#txaEditor_935")

            # 2. Update the Assunto paragraph
            # We select the <p> containing "Assunto:" inside the editor
            assunto_p = editor.locator("p:has-text('Assunto:')")
            assunto_p.fill("Assunto: Formulário de Empenho")

            # 3. Update the numbered paragraphs (Paragrafo_Numerado_Nivel1)
            paragraphs = editor.locator("p.Paragrafo_Numerado_Nivel1")

            paragraphs.nth(0).fill("1. Trata-se de solicitação de emissão de Nota de Empenho.")
            paragraphs.nth(1).fill(
                "2. Esta Subsecretaria de Assuntos Administrativos toma ciência da solicitação e encaminha os autos à Coordenação-Geral de Orçamento, Finanças e Contabilidade – CGOFC, para adoção das providências cabíveis."
            )

            # If there is a 3rd numbered paragraph leftover from the template, clear it:
            if paragraphs.count() > 2:
                paragraphs.nth(2).fill("")

            # 4. Update the centered signature lines directly using their specific SEI class
            # This keeps SEI's native centering class (Texto_Centralizado_EspacamentoSimples) intact!
            centered_lines = editor.locator("p.Texto_Centralizado_EspacamentoSimples")
     
                             
            centered_lines.nth(0).fill('GERÊNCIO NELCYR DE BEM')#("ANDREI RODRIGUES")
            centered_lines.nth(1).fill('Subsecretário de Assuntos Administrativos  - Substituto')#("Subsecretário de Assuntos Administrativos")
            locator.blur()  # Trigger change event

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"❌ Erro em access_proposal_details: {str(e)[:100]}")
            print(f"   Line: {exc_tb.tb_lineno}")
            return False

        time.sleep(0.3)
        
        # Step 6: Click Salvar button - use the first button in the editor toolbar
        try:
            # Try CKEditor's built-in save button in the toolbar
            save_clicked = new_page.evaluate('''
                () => {
                    // Try to find the Salvar button in CKEditor toolbar
                    const buttons = document.querySelectorAll('.ck-button');
                    for (let btn of buttons) {
                        const label = btn.getAttribute('aria-labelledby');
                        if (label) {
                            const labelEl = document.getElementById(label);
                            if (labelEl && labelEl.textContent.includes('Salvar')) {
                                btn.click();
                                return true;
                            }
                        }
                        // Also check title/attribute
                        const title = btn.getAttribute('data-cke-tooltip-text') || '';
                        if (title.includes('Salvar')) {
                            btn.click();
                            return true;
                        }
                    }
                    return false;
                }
            ''')
            
            if not save_clicked:
                # Try keyboard shortcut Ctrl+Alt+S
                new_page.keyboard.press('Control+Alt+s')
                print("⌨️ Atalho Ctrl+Alt+S pressionado")
            else:
                print("✅ Botão Salvar clicado")
            
            time.sleep(0.5)
            
        except Exception as e:
            print(f"⚠️ Erro ao salvar despacho: {str(e)[:80]}")
            # Try keyboard shortcut as fallback
            try:
                new_page.keyboard.press('Control+Alt+s')
                time.sleep(0.5)
            except:
                pass
        
        # Step 8: Close the editor window
        try:
            new_page.close()
            print("✅ Janela do editor fechada")
        except:
            pass
        
        # Switch back to main page
        for p in context.pages:
            if 'sei.mds.gov' in p.url and p != new_page:
                page = p
                break
        
        time.sleep(0.3)
        return True
    
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"❌ Erro em criar_despacho: {str(e)[:100]}")
        print(f"   Line: {exc_tb.tb_lineno}")
        return False


# ─── FUNÇÃO 4: Encaminhar para bloco de assinatura ───────────────────────────

def encaminhar_bloco_assinatura(page) -> bool:
    """
    Encaminhar processo para bloco de assinatura (432067 - ANDREI RODRIGUES).
    
    Flow:
    1. Click on "Incluir em Bloco de Assinatura" icon
    2. Select dropdown option 432067
    3. Click "Incluir"
    """
    print("📋 Encaminhando para bloco de assinatura (432067)...")
    
    try:
        # Step 1: Find and click the "Incluir em Bloco de Assinatura" link
        iframe_visualizacao = page.frame(name='ifrConteudoVisualizacao')
        btn_incluir = iframe_visualizacao.get_by_title("Incluir em Bloco de Assinatura")
        if btn_incluir:
            btn_incluir.click()
            time.sleep(0.3)

        else:
            # Try to find by SVG pattern
            page.evaluate('''
                () => {
                    const imgs = document.querySelectorAll('img');
                    for (let img of imgs) {
                        const alt = img.getAttribute('alt') || '';
                        const title = img.getAttribute('title') || '';
                        if (alt.includes('Bloco') || title.includes('Bloco')) {
                            img.click();
                            return;
                        }
                    }
                }
            ''')
            time.sleep(0.3)
        
        # Step 2: Wait for the form and select dropdown
        try:
            select_frame = page.wait_for_selector('#selBloco', timeout=5000)
            if select_frame:
                page.select_option('#selBloco', BLOCO_ASSINATURA_RP)
                print(f"✅ Bloco {BLOCO_ASSINATURA_RP} selecionado")
                time.sleep(0.5)
        except:
            # Try within iframe
            print("⚠️ Procurando selBloco em iframes...")
            for frame in page.frames:
                try:
                    if frame.query_selector('#selBloco'):
                        frame.select_option('#selBloco', BLOCO_ASSINATURA_RP)
                        print(f"✅ Bloco {BLOCO_ASSINATURA_RP} selecionado em iframe")
                        time.sleep(0.5)
                        break
                except:
                    continue
        
        # Step 3: Click "Incluir"
        try:
            incluir_btn = page.wait_for_selector('#sbmIncluir', timeout=5000)
            if incluir_btn:
                incluir_btn.click()
                print("✅ Processo incluído no bloco de assinatura")
                time.sleep(0.3)
        except:
            # Try within iframe
            for frame in page.frames:
                try:
                    if frame.query_selector('#sbmIncluir'):
                        frame.click('#sbmIncluir')
                        print("✅ Processo incluído no bloco de assinatura")
                        time.sleep(0.3)
                        break
                except:
                    continue
        
        return True
    
    except Exception as e:
        print(f"⚠️ Erro ao encaminhar para bloco de assinatura: {str(e)[:80]}")
        return False


# ─── FUNÇÃO 5: Encaminhar para bloco interno ─────────────────────────────────

def encaminhar_bloco_interno(page, context) -> bool:
    """
    Encaminhar processo para bloco interno (434283 - RP e GRU pelo Robôzinho).
    
    Flow:
    1. In ifrArvore, click on the process number link (second anchor)
    2. Click "Incluir em Bloco de Assinatura" again
    3. In the modal/popup, select bloco interno 434283
    """
    print("📋 Encaminhando para bloco interno (434283)...")
    
    try:
        # Step 1: In ifrArvore, click the process number link
        frame_arvore = page.frame(name='ifrArvore')
        if frame_arvore:
            try:
                # Find the span with processoNaoVisualizado or processoVisualizado
                frame_arvore.evaluate('''
                    () => {
                        const anchors = document.querySelectorAll('#divArvore a');
                        // Find the second anchor that contains a process number (has span with number)
                        let count = 0;
                        for (let a of anchors) {
                            const span = a.querySelector('span');
                            if (span && span.textContent.includes('.')) {
                                count++;
                                if (count === 2) {
                                    a.click();
                                    return;
                                }
                            }
                        }
                    }
                ''')
                time.sleep(0.3)
            except:
                pass
        
        # Step 2: Click "Incluir em Bloco de Assinatura" again
        # This should open a modal/popup
        bloco_link = page.query_selector('img[alt="Incluir em Bloco de Assinatura"], img[title*="Bloco de Assinatura"]')
        if bloco_link:
            bloco_link.click()
            time.sleep(0.3)
        else:
            page.evaluate('''
                () => {
                    const imgs = document.querySelectorAll('img');
                    for (let img of imgs) {
                        const alt = img.getAttribute('alt') || '';
                        const title = img.getAttribute('title') || '';
                        if (alt.includes('Bloco') || title.includes('Bloco')) {
                            img.click();
                            return;
                        }
                    }
                }
            ''')
            time.sleep(0.3)
        
        # Step 3: The modal might be a new page or an iframe
        # Check for modal iframe
        modal_page = None
        
        # Check if a new page opened
        for p in context.pages:
            if 'bloco_selecionar_processo' in p.url:
                modal_page = p
                break
        
        if modal_page:
            # Working in a new page
            print("✅ Modal de bloco interno detectado (nova página)")
            
            # Select the radio for 434283
            try:
                modal_page.evaluate(f'''
                    () => {{
                        const radios = document.querySelectorAll('input[type="radio"]');
                        for (let radio of radios) {{
                            if (radio.value === "{BLOCO_INTERNO_RP}") {{
                                radio.click();
                                break;
                            }}
                        }}
                    }}
                ''')
                time.sleep(0.5)
                
                # Click the transport link
                modal_page.evaluate(f'''
                    () => {{
                        const links = document.querySelectorAll('a[id*="lnkInfraT"]');
                        for (let link of links) {{
                            if (link.id === "lnkInfraT-{BLOCO_INTERNO_RP}") {{
                                link.click();
                                break;
                            }}
                        }}
                    }}
                ''')
                print(f"✅ Bloco interno {BLOCO_INTERNO_RP} selecionado")
                time.sleep(0.3)
                
            except Exception as e:
                print(f"⚠️ Erro no modal de bloco interno: {str(e)[:80]}")
        
        else:
            # Try to find modal in iframes
            for frame in page.frames:
                try:
                    radio = frame.query_selector(f'input[value="{BLOCO_INTERNO_RP}"]')
                    if radio:
                        radio.click()
                        time.sleep(0.5)
                        
                        # Click transport link
                        transport_link = frame.query_selector(f'#lnkInfraT-{BLOCO_INTERNO_RP}')
                        if transport_link:
                            transport_link.click()
                            print(f"✅ Bloco interno {BLOCO_INTERNO_RP} selecionado")
                            time.sleep(0.3)
                        break
                except:
                    continue
        
        return True
    
    except Exception as e:
        print(f"⚠️ Erro ao encaminhar para bloco interno: {str(e)[:80]}")
        return False


# ─── FUNÇÃO 6: Limpar e voltar para controle de processos ─────────────────────

def voltar_controle_processos(page):
    """Return to Controle de Processos page and refresh"""
    try:
        # Try to click Controle de Processos link
        page.click('#lnkControleProcessos')
        time.sleep(0.3)
    except:
        try:
            # Try navigation
            page.goto(page.url.split('?')[0])
            time.sleep(0.3)
        except:
            # Force refresh
            page.reload()
            time.sleep(0.3)
    
    wait_for_frames_to_load(page, timeout=10000)
    print("🔄 Voltou para Controle de Processos")


# ─── FUNÇÃO PRINCIPAL ────────────────────────────────────────────────────────

def executar_scraping():
    """Main execution function - processes new SEI processes"""
    
    other_door = '9232'
    
    page = None
    playwright = None
    browser = None
    
    try:
        if isinstance(other_door, str) and other_door.isdigit():
            page, playwright, browser = conectar_navegador_existente(int(other_door))
        
        if not page:
            print("❌ Failed to connect to browser. Exiting.")
            return
        
        context = page.context
        
        while True:
            try:
                # Step 1: Extract new processes (non-visualized / red)
                #processos_novos = extract_received_processos(page=page)
                processos_novos = ['71000.059301/2026-35']
                
                if processos_novos:
                    print(f"\n🔄 Processando {len(processos_novos)} processos novos...")
                    
                    for i, processo in enumerate(sorted(processos_novos), 1):
                        print(f"\n{'='*60}")
                        print(f"📌 Processo {i}/{len(processos_novos)}: {processo}")
                        print(f"{'='*60}")
                        
                        # Step 2: Access and check if process has RP + SNEAELIS
                        if access_proposal_details(page=page, numero_processo=processo):
                            try:  
                                # Step 3: Create despacho
                                if criar_despacho(page=page, context=context, numero_processo=processo):

                                    print(f"✅ Despacho criado para processo {processo}")
                                    # Step 4: Encaminhar para bloco de assinatura
                                    encaminhar_bloco_assinatura(page=page)
                                    
                                    # Step 5: Encaminhar para bloco interno
                                    encaminhar_bloco_interno(page=page, context=context)
                                    
                                    print(f"✅ Processo {processo} concluído com sucesso!")
                            except Exception as e:
                                    print(f"⚠️ Falha ao criar despacho para {processo}\nERRO: type={type(e).__name__}\n msg={str(e)[:100]}")
                    else:
                        print(f"⏭️ Processo {processo} não atende critérios, pulando")
                    
                    # Back to Controle de Processos
                    voltar_controle_processos(page)
                else:
                    print("⏳ Nenhum processo novo encontrado. Aguardando...")
                
                # Refresh page
                print("\n🔄 Atualizando página...")
                page.reload()
                wait_for_frames_to_load(page, timeout=10000)

                sys.exit()
                # Delay before next iteration
                #time.sleep(300)
            
            except KeyboardInterrupt:
                raise
            except Exception as e:
                print(f"❌ Erro no loop principal: {type(e).__name__}")
                print(f"   Mensagem: {str(e)[:100]}")
                traceback.print_exc()
                time.sleep(0.3)
    
    except KeyboardInterrupt:
        print("\n🛑 Script interrompido pelo usuário")
    except Exception as e:
        print(f"❌ Erro fatal: {type(e).__name__}")
        print(f"   Mensagem: {str(e)[:100]}")
        traceback.print_exc()
    finally:
        if playwright:
            playwright.stop()
        print("✅ Recursos do Playwright liberados")


# ─── EXECUÇÃO ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    start_time = time.perf_counter()
    executar_scraping()
    elapsed = time.perf_counter() - start_time
    print(f"⏱️ Tempo total de execução: {elapsed:.2f} segundos")