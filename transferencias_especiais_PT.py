from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (TimeoutException, NoSuchElementException,
                                        ElementClickInterceptedException, WebDriverException,
                                        StaleElementReferenceException, ElementNotInteractableException)
from selenium.webdriver.chrome.service import Service
import pandas as pd
import time
# Função para conectar ao navegador já aberto
def conectar_navegador_existente():
    """
    Inicializa o Objeto Robo, configurando e iniciando o driver do Chrome.
    """
    try:
        # Configuração do registro
        #  .arquivo_registro = ''
        # Inicia as opções do Chrome
        chrome_options = webdriver.ChromeOptions()
        # Endereço de depuração para conexão com o Chrome
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        # Inicializa o driver do Chrome com as opções e o gerenciador de drivers
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options= chrome_options)
        driver.switch_to.window( driver.window_handles[0])
        print("✅ Conectado ao navegador existente com sucesso.")    

        return driver
    except WebDriverException as e:
        # Imprime mensagem de erro se a conexão falhar
        print(f"❌ Erro ao conectar ao navegador existente: {e}")

# Trunca a mensagem de erro
def truncate_error(msg, max_error_length=100):
    """Truncates error message with ellipsis if too long"""
    return (msg[:max_error_length] + '...') if len(msg) > max_error_length else msg

# Função para clicar em um elemento com retry
def clicar_elemento(driver, xpath, retries=3):
    """Attempts to click an element with truncated error messages.

    Args:
        driver: WebDriver instance
        xpath: XPath of the element to click
        retries: Number of retry attempts
    """
    last_error = None

    for tentativa in range(1, retries + 1):
        try:
            remover_backdrop(driver)
            elemento = WebDriverWait(driver, 7).until(
                EC.element_to_be_clickable((By.XPATH, xpath)))
            print(f"✅ Clicked element (attempt {tentativa}): {elemento.text[:30]}...")

            if elemento.is_displayed():
                elemento.click()
                return True
        except StaleElementReferenceException:
            print("🔄 Stale element detected! Attempting recovery...")  # Debug print
            # Wait for the old element to go stale, then re-find
            WebDriverWait(driver, 10).until(EC.staleness_of(elemento))
            new_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, xpath)))
            new_element.click()

        except ElementClickInterceptedException as e:
            last_error = truncate_error(f"Element intercepted on click (attempt {tentativa}): {str(e)}")
            print(f"⚠️ {last_error}")

        except TimeoutException as e:
            last_error = truncate_error(f"Timeout (attempt {tentativa}): {str(e)}")
            print(f"⏱️ {last_error}")

        except NoSuchElementException as e:
            last_error = truncate_error(f"Element not found (attempt {tentativa}): {str(e)}")
            print(f"🔍 {last_error}")

        except Exception as e:
            last_error = truncate_error(f"Unexpected error (attempt {tentativa}): {str(e)}")
            print(f"❌ {last_error}")

        if tentativa < retries:
            time.sleep(2 * tentativa)

    print(truncate_error(f"❌ Failed after {retries} attempts for {elemento.text}. Last error: {last_error}"))
    return False

# Função para inserir texto em um campo com retry
def inserir_texto(driver, xpath, texto, retries=3):
    for tentativa in range(retries):
        try:
            remover_backdrop(driver)  # Remove o backdrop antes de inserir texto
            elemento = WebDriverWait(driver, 7).until(
                EC.presence_of_element_located((By.XPATH, xpath)))
            elemento.clear()
            elemento.send_keys(texto)
            print(f"✅ Texto inserido no campo:{elemento.text}")
            return True
        except Exception as erro:
            print(f"❌ Erro ao inserir texto no campo {xpath} (tentativa {tentativa + 1}): {erro}")
            time.sleep(1)  # Espera antes de tentar novamente
    print(f"❌ Falha após {retries} tentativas para inserir texto em {xpath}")
    return False

# Função para obter texto de um elemento
def obter_texto(driver, xpath):
    try:
        remover_backdrop(driver)  # Remove o backdrop antes de obter texto
        elemento = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, xpath)))
        texto = elemento.text.strip()
        return texto if texto else None
    except Exception as erro:
        print(f"❌ Erro ao obter texto do elemento {xpath}: {erro}")
        return None

# Evidencia o elemento que está sendo tocado
def highlight_element(driver, element, effect_time=3, color="red", border=3):
    """Highlights an element with a colored border"""
    driver.execute_script(
        """
        let oldStyle = arguments[0].getAttribute('style');
        arguments[0].setAttribute('style', 
            `border: {border}px solid {color}; 
             box-shadow: 0px 0px 10px {color}; 
             background: rgba(255,255,0,0.2); 
             transition: all 0.3s ease;`);
        setTimeout(function() {{
            arguments[0].setAttribute('style', oldStyle || '');
        }}, {effect_time} * 1000);
        """.format(border=border, color=color, effect_time=effect_time),
        element
    )

# Verifica o valor dos campos com seletores binários
def get_radio_selection(driver, xpath):
    def get_angular_radio_state():
        """
        Robustly gets radio button state in Angular applications with:
        - Angular stability waiting
        - Multiple fallback methods
        - Comprehensive error handling
        - Automatic retries
        """

        max_attempts = 3
        for attempt in range(max_attempts):
            try:
                # 1. Wait for Angular to stabilize (critical for Angular apps)
                print(f"Attempt {attempt + 1}: Waiting for Angular stability...")
                wait_for_angular()

                # 2. Wait for the specific element to be interactable
                radio_group = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, 'br-radio[formcontrolname="isRecursosIndicados"]'))
                )

                # 3. Try multiple detection methods with fallbacks
                print("Trying detection methods...")
                value = execute_with_fallbacks(radio_group)

                if value is not None:
                    return value

                print("Retrying...")
                time.sleep(1)  # Brief pause before retry

            except Exception as e:
                print(f"Attempt {attempt + 1} failed:", str(e))

        print("All attempts exhausted. Saving debug information...")
        return None

    def wait_for_angular():
        """Specialized wait for Angular applications"""
        try:
            driver.execute_async_script("""
            var callback = arguments[arguments.length - 1];
            if (window.getAllAngularTestabilities) {
                var testabilities = window.getAllAngularTestabilities();
                var count = testabilities.length;
                if (count === 0) return callback(true);

                var decrement = function() {
                    count--;
                    if (count === 0) callback(true);
                };

                testabilities.forEach(function(t) {
                    t.whenStable(decrement);
                });
            } else {
                callback(true);  // Proceed if not Angular app
            }
            """)
        except:
            # Fallback to simple wait if Angular detection fails
            time.sleep(1)

    def execute_with_fallbacks(radio_group):
        """Tries multiple detection methods in priority order"""
        methods = [
            # 1. Check native DOM properties first
            lambda: driver.execute_script("""
                const input = arguments[0].querySelector('input[type="radio"]:checked');
                return input ? input.value : null;
            """, radio_group),

            # 2. Check Angular model binding
            lambda: radio_group.get_attribute('ng-reflect-model'),

            # 3. Check for hidden form control value
            lambda: driver.execute_script("""
                const form = arguments[0].closest('form');
                if (!form) return null;
                const formCtrl = window.ng?.getComponent(form)?.form?.get('isRecursosIndicados');
                return formCtrl?.value ?? null;
            """, radio_group),

            # 4. Check visual state as last resort
            lambda: check_visual_state(radio_group)
        ]

        for method in methods:
            try:
                result = method()
                if result is not None:
                    print(f"Method succeeded: {result}")
                    return str(result).lower()  # Normalize to string
            except:
                continue

        return None

    def check_visual_state( element):
        """Fallback for visually determining state"""
        active_class = driver.execute_script("""
        return window.getComputedStyle(arguments[0]).getPropertyValue('active');
        """, element)
        return 'true' if active_class and 'active' in active_class else None

    try:
        radio_button = get_angular_radio_state()

        # Execute the script and get the result
        selected_value = radio_button
        print(selected_value)

        # Determine the selected option
        if selected_value:
            print("🔘 Selected: Sim (true)")
            return "Sim"
        else:
            print("🔘 Selected: Não (false)")
            return "Não"

    except Exception as e:
        print(f"⚠️ Error checking radio buttons: {str(e)[:60]}")
        print(f"🔘 Opção selecionada: Não")
        return "Não"


# Função para obter o valor de um campo desabilitado
def obter_valor_campo_desabilitado(driver, xpath, call=None):
    try:
        remover_backdrop(driver)  # Remove o backdrop antes de obter valor
        try:
            elemento = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, xpath)))
        except TimeoutException:
            print(f"⏱️ Timeout: Elemento não encontrado em 5s - {truncate_error(xpath)}")
            return "Campo Vazio"
        except NoSuchElementException:
            print(f"🔍 Elemento não existe - {truncate_error(xpath)}")
            return "Campo Vazio"

            # 2. Tentar métodos alternativos para obter valor
        valor = None

        # Métod 1: Atributo value padrão
        try:
            WebDriverWait(driver, 2).until(
                lambda d: elemento.get_attribute("value") and elemento.get_attribute("value").strip() != "")
            valor = elemento.get_attribute("value")
        except TimeoutException:
            pass  # Vamos tentar outros métodos

        # Métod 2: JavaScript para campos desabilitados
        if not valor:
            try:
                valor = driver.execute_script(
                    "return arguments[0].value || arguments[0].defaultValue;",
                    elemento)
            except Exception as js_err:
                print(f"⚠️ JS fallback falhou: {truncate_error(str(js_err))}")
                pass

        # Métod 3: JavaScript para campos desabilitados
        if not valor:
            try:
                valor = elemento.text
            except AttributeError:
                valor = None  # or use a default value like '' or 'Não encontrado'
                print("⚠️ Aviso: elemento não possui atributo 'text'.")
            except Exception as e:
                valor = None
                if call:
                    print(f'Erro gerado pela call nº{call}')
                print(f"❌ Erro inesperado ao acessar elemento.text: {type(e).__name__} - {e}")

        # Tratamento do valor retornado
        if valor and str(valor).strip():
            valor = str(valor).strip()
            print(f"📦📄 Valor obtido: {valor[:50]}...")  # Trunca valores longos
            return valor
        else:
            print(f"ℹ️ Campo:{elemento.text} vazio ou sem valor válido")
            return
    except StaleElementReferenceException:
        print(f"👻 Elemento tornou-se obsoleto - {truncate_error(xpath)}")
        return
    except Exception as erro:
        print(f"❌ Erro ao tentar obetar valor do campo. Err: {type(erro).__name__} -"
              f" {truncate_error(str(erro))}")
        return "Campo Vazio"

# Função para remover o backdrop
def remover_backdrop(driver):
    try:
        driver.execute_script("""
            var backdrops = document.getElementsByClassName('backdrop');
            while(backdrops.length > 0){
                backdrops[0].parentNode.removeChild(backdrops[0]);
            }
        """)
        #print("✅ Backdrop removido com sucesso!")
    except Exception as erro:
        print(f"❌ Erro ao remover backdrop: {erro}")

# Função que aplica a espera pelo elemento
def wait_for_element(driver, xpath: str, timeout: int = 10) -> bool:
    """Aguarda até que um elemento esteja presente na página."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath)))
        return True
    except Exception:
        return False

 # Extrai os dados da tabela de forma concatenada
def extract_all_rows_text(driver, table_xpath, timeout=10, count: int=0):
    """
    Extracts all text content from each row in an ngx-datatable.

    Args:
        driver: Selenium WebDriver instance
        table_xpath: XPath of the datatable
        timeout: Maximum wait time in seconds

    Returns:
        list: List of strings containing each row's full text content
        or None if failed
    """
    try:
        # Wait for at least one row to be present
        table = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, table_xpath)))

        # Now fetch all rows
        rows = table.find_elements(By.CSS_SELECTOR, 'datatable-body-row')

        row_data = []
        if not rows:
            print("⚠️ No rows found in table")
            return None

        for idx, row in enumerate(rows, start=1):
            # For each row, find all the cells
            cells = row.find_elements(By.CSS_SELECTOR, 'datatable-body-cell')

            for cell in cells:
                # Each cell has a label div where the text is
                label = cell.find_element(By.CSS_SELECTOR, 'div.datatable-body-cell-label')
                text = label.text.strip()
                row_data.append(text)

        joined_row = " | ".join(row_data)

        return joined_row

    except TimeoutException:
        print(f"⌛ Timeout: No rows appeared within {timeout} seconds")
        return None
    except Exception as e:
        print(f"❌ Extraction failed: {str(e)[:100]}...")
        return None

# Extrai os dados da Lista de Documentos Hábeis
def extrat_all_cells_text(driver, table_xpath, col_idx, timeout=10):
    """
    Extracts all text content from a specific column in a ngx-datatable across all pages.

    Args:
        driver: Selenium WebDriver instance
        table_xpath: XPath of the datatable
        col_idx: Index of the column to extract (0-based)
        timeout: Maximum wait time in seconds

    Returns:
        str: All column data joined with " | " or None if failed
    """
    col_data = []
    processed_pages = set()
    current_page = 1

    try:
        while True:
            # Wait for table to be present
            table = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, table_xpath))
            )

            # Get all rows in current page
            rows = table.find_elements(By.CSS_SELECTOR, 'datatable-body-row')
            if not rows:
                print("⚠️ No rows found in table")
                break
            # Get current page information (e.g., current page number text or unique identifier)
            try:
                pagination_info = driver.find_element(By.XPATH,
                                                  '/html/body/transferencia-especial-root/br-main-layout/div/div'
                                                  '/div/main/transferencia-especial-main/transferencia-plano-acao'
                                                  '/transferencia-cadastro/br-tab-set/div/nav'
                                                  '/transferencia-plano-acao-dados-orcamentarios/br-table[2]/div'
                                                  '/ngx-datatable/div/datatable-footer/div/br-pagination-table'
                                                  '/div/br-select[2]')
                pagination_info = pagination_info.text.strip()
            except NoSuchElementException:
                pagination_info = str(current_page)

            if f"{current_page}" in processed_pages:
                print(f"⚠️ Already processed page {current_page}, breaking loop")
                break

            processed_pages.add(current_page)

            for row in rows:
                cells = row.find_elements(By.CSS_SELECTOR, 'datatable-body-cell')
                if len(cells) > col_idx:
                    cell = cells[col_idx]
                    label = cell.find_element(By.CSS_SELECTOR, 'div.datatable-body-cell-label')
                    # Try to find span or button with text
                    text_elements = label.find_elements(By.XPATH, './/*[text()]')
                    if text_elements:
                        text = text_elements[0].text.strip()
                    else:
                        text = label.text.strip()
                    if text:  # Only add non-empty text
                        col_data.append(text)
            try:
                next_button = driver.find_element(By.XPATH,
                               '//button[contains(@class, "br-button") and contains(@class, "next")]')
                if not next_button.is_enabled():
                    print("ℹ️ No more pages to process")
                    break
                next_button.click()
                current_page += 1
                # Wait for page to load
                WebDriverWait(driver, timeout).until(
                    EC.staleness_of(rows[0])  # Wait until old rows are gone
                )
            except (NoSuchElementException, ElementNotInteractableException):
                print("ℹ️ No next page button found or not clickable")
                break

        return " | ".join(col_data) if col_data else None

    except Exception as e:
        print(f"❌ Error extracting table data: {str(e)}")
        return None


# Navega pela primeira página e coleta os dados
def loop_primeira_pagina(driver, plano_acao: dict):
    #  [0] Beneficiário (beneficiario['nome'])  //  [1] UF (beneficiario['uf'])  //
    #  [2] Banco (dados_bancarios['banco'])  //  [3] Agência (dados_bancarios['agencia'])  //
    #  [4] Conta (dados_bancarios['conta'])  //  [5] Situação Conta (dados_bancarios['situacao'])  //
    #  [6] Emenda (emenda['numero'])  //  [7] Valor Emenda (emenda['valor'])  //
    #  [8] Finalidade  //  [9] Programações Orçamentárias selecionadas  //
    lista_caminhos = [
        # Beneficiário [0][1]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[1]/fieldset/div[2]/div[1]/div[1]/br-input/div/div[1]/input',

        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[1]/fieldset/div[2]/div[1]/div[2]/br-input/div/div[1]/input',

        # Banco [2]>[5]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[1]/fieldset/div[2]/div[2]/div[1]/br-input/div/div[1]/input',

        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[1]/fieldset/div[2]/div[2]/div[2]/br-input/div/div[1]/input',

        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[1]/fieldset/div[2]/div[2]/div[3]/br-input/div/div[1]/input',

        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[1]/fieldset/div[2]/div[2]/div[4]/br-input/div/div[1]/input',


        # Dados emenda parlamentar [6][7]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[2]/fieldset/div[2]/div/div[1]/br-input/div/div[1]/input',

        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[2]/fieldset/div[2]/div/div[3]/br-input/div/div[1]/input',

        # Finalidade [8]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[3]/fieldset/div[2]',

        # Programações Orçamentárias selecionadas [9]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/'
        'transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/'
        'transferencia-plano-acao-dados-basicos/form/br-fieldset[3]/fieldset/div[2]/div[2]/div/br-table/div/'
        'ngx-datatable/div/datatable-body/datatable-selection/datatable-scroller/datatable-row-wrapper/'
        'datatable-body-row/div[2]/datatable-body-cell[1]/div'
    ]

    plano_acao["beneficiario"]["nome"] = obter_valor_campo_desabilitado(driver, lista_caminhos[0])
    plano_acao["beneficiario"]["uf"] = obter_valor_campo_desabilitado(driver, lista_caminhos[1])
    plano_acao["dados_bancarios"]["banco"] = obter_valor_campo_desabilitado(driver, lista_caminhos[2])

    plano_acao["dados_bancarios"]["agencia"] = obter_valor_campo_desabilitado(driver, lista_caminhos[3])
    plano_acao["dados_bancarios"]["conta"] = obter_valor_campo_desabilitado(driver, lista_caminhos[4])
    plano_acao["dados_bancarios"]["situacao"] = obter_valor_campo_desabilitado(driver, lista_caminhos[5])

    plano_acao["emenda"]["numero"] = obter_valor_campo_desabilitado(driver, lista_caminhos[6])
    plano_acao["emenda"]["valor"] = obter_valor_campo_desabilitado(driver, lista_caminhos[7])
    plano_acao["finalidade"] = extract_all_rows_text(driver=driver, table_xpath=lista_caminhos[8])

    plano_acao["programacoes_orcamentarias"] = extract_all_rows_text(driver, lista_caminhos[9])

    return plano_acao

# Navega pela primeira página e coleta os dados
def loop_segunda_pagina(driver,index, plano_acao: dict, df, df_path):
    # [0] Aba Dados Orçamentários // [1] Pagamentos Empenho // [2] Pagamentos Valor //
    # [3] Pagamentos Ordem // [4] Aba Plano de Trabalho // [5] Declaracoes Recursos Orcamento //
    # [6] Declaracoes Nao Uso Pessoal // [7] Execucao Executor // [8] Execucao Objeto //
    # [9] Metas Descricao // [10] Metas Quantidade // [11] Metas Unidade //
    # [12] Metas Meses // [13] Historico Sistema // [14] Historico Concluido //
    # [15] Controle Social Última notificação // [16] Controle Social Resp // [17] Período de Execução
    # [18] Classificação Orçamentária de Despesa // [19] Prazo de Execução em meses
    lista_caminhos = [
        # [0]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/'
        'transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/ul/li[2]/button/span',

        # [1] > [3]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-orcamentarios/br-table[2]/div/ngx-datatable',

        # [2]
        '',

        # [3]
        '',

        # [4]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/'
        'transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/ul/li[3]/button/span',

        # [5]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-plano'
        '-trabalho/form/div/div/br-fieldset[1]/fieldset/div[2]/div[2]/div/br-radio/div/div[1]/input',

        # [6]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-plano'
        '-trabalho/form/div/div/br-fieldset[2]/fieldset/div[2]/div[1]/div/br-checkbox/div/label',

        # [7] > Lista onde os dados de [7] à [12]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-plano'
        '-trabalho/form/div/div/br-fieldset[3]/fieldset/div[2]/div[2]/div[2]/br-table/div/ngx-datatable/div',

        # [8] > [13][14]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-plano'
        '-trabalho/br-fieldset[2]/fieldset/div[2]/div/div/br-table/div/ngx-datatable',

        #[9] > [15][16]
        '',

        # [10] > [17]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-plano'
        '-trabalho/transferencia-plano-trabalho-resumo/div/div[6]/div[2]/div',
        
        # [11] > [18]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-plano'
        '-trabalho/form/div/div/br-fieldset[1]/fieldset/div[2]/div[3]/div/br-textarea/div/div['
        '1]/div/textarea',

        # [12] > [19]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao'
        '-plano-trabalho/form/div/div/br-fieldset[2]/fieldset/div[2]/div[2]/div[1]/br-input/div/div['
        '1]/input'
    ]

    try:
        # Aba Dados Orçamentários
        clicar_elemento(driver, lista_caminhos[0])
        time.sleep(0.5)
        # Empenho section
        plano_acao["pagamentos"]["empenho"] = extrat_all_cells_text(driver, lista_caminhos[1], 0)
        plano_acao["pagamentos"]["valor"] = extrat_all_cells_text(driver, lista_caminhos[1], 3)
        plano_acao["pagamentos"]["ordem"] = extrat_all_cells_text(driver, lista_caminhos[1], 5)
        # Aba Plano de Trabalho
        clicar_elemento(driver, lista_caminhos[4])

        # Declarações section
        time.sleep(0.5)
        plano_acao["declaracoes"]["recursos_orcamento"] = get_radio_selection(driver, lista_caminhos[5])
        plano_acao["declaracoes"]["nao_uso_pessoal"] = obter_valor_campo_desabilitado(driver,
                                                                                      lista_caminhos[6])
        if plano_acao["declaracoes"]["nao_uso_pessoal"] not in ['', 'Campo Vazio']:
            plano_acao["declaracoes"]["nao_uso_pessoal"] = 'Sim'
        else:
            plano_acao["declaracoes"]["nao_uso_pessoal"] = 'Não'

        # Classificação Orçamentária de Despesa
        plano_acao["classificacao_orcamentaria"] = obter_valor_campo_desabilitado(driver, lista_caminhos[11])

        # Prazo de Execução em meses
        plano_acao["prazo_de_execucao"] = obter_valor_campo_desabilitado(driver, lista_caminhos[12])

        # Execução and Metas section
        coletar_dados_listas(driver, lista_caminhos[7], index=index, df=df, df_path=df_path)

        # Histórico section
        coletar_dados_hist(driver, lista_caminhos[8], index=index, df_path=df_path)

        # Controle Social section
        plano_acao["controle_social"]["conselhos"] = obter_valor_campo_desabilitado(driver, lista_caminhos[9])
        plano_acao["controle_social"]["instancias"] = obter_valor_campo_desabilitado(driver, lista_caminhos[9])

        # Periodo_exec
        plano_acao["periodo_exec"] = obter_valor_campo_desabilitado(driver, lista_caminhos[10])

        return plano_acao

    except Exception as erro:
        print(f"❌ Erro inesperado: {type(erro).__name__} - {truncate_error(str(erro))}")
        return None

# Coleta os dados do histórico
def coletar_dados_hist(driver, tabela_xpath, index, df_path):
    # 2. Helper function to check empty rows
    def is_row_empty(row_idx):
        try:
            return all(
                pd.isna(df.at[row_idx, col]) or
                str(df.at[row_idx, col]).strip() in ('', 'None', 'nan')
                for col in colunas_para_salvar
            )
        except KeyError as err:
            print(f"⚠️ Missing column {err}")
            return True

    def is_first_cell_empty(row_idx: int, col_id: int=0) -> bool:
        """
        Checks whether the cell at the given row and column index is empty or NaN.

        Parameters:
            df (pd.DataFrame): The DataFrame to check.
            row_idx (int): Index of the row.
            col_idx (int): Index of the column (e.g., 0 for the first column).

        Returns:
            bool: True if the cell is empty or NaN, or if the row is out of bounds.
        """
        if row_idx >= len(df) or col_id >= len(df.columns):
            return True
        cell = df.iat[row_idx, col_id]
        return pd.isna(cell) or cell == ""

    print(f"🚀 Starting coletar_dados_hist() for index {index}")

    # Read data frame
    df = pd.read_excel(df_path, dtype=str)

    colunas_para_salvar = ["Responsável", "Data", "Situação"]
    # Initialize data variables
    sys_data = ["Não Encontrado"] * 3
    conc_data = ["Não Encontrado"] * 3
    table_data = []

    try:
        # Wait for table to be present
        tabela = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, tabela_xpath))
        )
        print("🔎📋 Tabela de histórico localizada")
        try:
            # Get all rows in the table body
            rows = tabela.find_elements(By.XPATH, f"{tabela_xpath}//datatable-body-row")
        except Exception as e:
            print(f"⚠️ Error finding row. Err: {str(e)[:50]}")
            table_data = ["Não Encontrado"] * 6
            return table_data

        # Process each row to extract cell data
        for i, row in enumerate(rows):
            try:
                cells = row.find_elements(By.XPATH, ".//datatable-body-cell")
                row_data = [cell.text.strip() if cell.text else "Empty" for cell in cells]
                table_data.append(row_data)
            except Exception as e:
                print(f"⚠️ Error processing row {i}: {str(e)[:50]}")
                continue

        # Search backwards for "Sistema" row
        for k in range(len(table_data) - 1, -1, -1):
            try:
                if table_data[k][0] == "Sistema":
                    sys_data = table_data[k][:3]  # First 3 columns

                    # Look for "Concluído" in subsequent rows
                    for j in range(k + 1, len(table_data)):
                        if table_data[j][2] == "Concluído":
                            conc_data = table_data[j][:3]
                            break
            except Exception as e:
                print(f"❌ Falha na extração de dados histórico: {type(e).__name__} -"
                      f" {truncate_error(str(e))}")

        try:
            empty_row = next((i for i in range(index, len(df) + 1) if is_row_empty(i)), len(df))
            print(empty_row)
            # 3. Save Sistema data
            if is_row_empty(index):
                print(f"💾 Saving Sistema data to row {index}")
                for col, val in zip(colunas_para_salvar, sys_data):
                    df.at[index, col] = val
                    #print(df.iloc[index, [19, 20, 21, 22, 23, 24, 25, 26]])
            else:
                print(f"⚠️ Row {index} not empty. Finding next available...")
                if empty_row >= len(df):
                    df.loc[empty_row] = [None] * len(df.columns)
                for col, val in zip(colunas_para_salvar, sys_data):
                    df.at[empty_row, col] = val
                print(f"💾 Saved Sistema data to row {empty_row}")

            # 4. Save Concluído data
            conc_start = index + 1
            current_cell = is_first_cell_empty(conc_start, 24)
            first_cell = is_first_cell_empty(conc_start)

            nova_linha_data = {col:conc_data[q] for q, col in enumerate(colunas_para_salvar) }

            try:
                if current_cell and first_cell:
                    # Insert data in current row
                    for col_idx, col in enumerate(colunas_para_salvar):
                        df.at[conc_start, col] = conc_data[col_idx]
                    print(f"ℹ️ Dados inseridos na linha existente {conc_start}")

                elif not first_cell:
                    # Current cell has data, insert new row
                    new_row = {col: nova_linha_data.get(col, '') for col in df.columns}
                    new_row_df = pd.DataFrame([new_row])
                    df = pd.concat([df.iloc[:conc_start], new_row_df, df.iloc[conc_start:]],
                                   ignore_index=True)

                    print(f"ℹ️ Nova linha criada na posição {conc_start}")
                else:
                    conc_start += 1

                    # Convert the new data to a DataFrame
                    new_row = {col: nova_linha_data.get(col, '') for col in df.columns}
                    new_row_df = pd.DataFrame([new_row])

                    # Split the original DataFrame
                    upper = df.iloc[:conc_start]
                    lower = df.iloc[conc_start:]

                    df = pd.concat([upper, new_row_df, lower], ignore_index=True)

                    print(f"ℹ️ Nova linha criada na posição {conc_start + 1}")
            except Exception as erro:
                print(f"❌ Erro ao salvar listas no Excel: {type(erro).__name__} -"
                      f" {truncate_error(str(erro))}")
        except Exception as e:
            print(f"❌ Erro ao organizar DataFrame: {type(e).__name__} - {truncate_error(str(e))}")

        try:
            # Save the DataFrame
            df.to_excel(df_path, index=False, engine='openpyxl')
        except Exception as e:
            print(f"❌ Erro ao salvar dados histórico: {type(e).__name__} - {truncate_error(str(e))}")

    except Exception as erro:
        print(f"❌ Erro ao coletar dados histórico: {type(erro).__name__} - {truncate_error(str(erro))}")

# Função para coletar executores e metas
def coletar_dados_listas(driver, tabela_xpath, index, df, df_path):
    print(f"🚀 Starting coletar_dados_listas() for index {index}")
    # Extracts all metas from the currently expanded executor row
    def extract_metas_from_expanded_row():
        """
        Extracts all metas from the currently expanded executor row.
        Returns a list of dictionaries: each with meta, unidade_medida, quantidade.
        """

        def _wait_for_angular_stable():
            """Helper to wait for Angular to finish rendering"""
            driver.execute_async_script("""
            var callback = arguments[arguments.length - 1];
            if (window.getAllAngularTestabilities) {
                var testabilities = window.getAllAngularTestabilities();
                var count = testabilities.length;
                var decrement = function() {
                    count--;
                    if (count === 0) callback(true);
                };
                testabilities.forEach(function(t) {
                    t.whenStable(decrement);
                });
            } else {
                callback(true);
            }
            """)

        try:
            # 1. Wait for Angular to stabilize
            _wait_for_angular_stable()

            # 2. Find expanded container with multiple fallback strategies
            expanded_container = None
            locators = [
                # Primary locator
                (By.XPATH,
                 "//div[contains(@class, 'datatable-row-detail') and not(contains(@style, 'none'))]"),
                # Fallback locator 1
                (By.CSS_SELECTOR, "div.datatable-row-detail:not([style*='display: none'])"),
                # Fallback locator 2
                (By.XPATH, "//div[contains(@class, 'expanded-container')]")
            ]

            for locator in locators:
                try:
                    expanded_container = WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located(locator)
                    )
                    break
                except:
                    continue

            if not expanded_container:
                raise Exception("Could not find expanded row container")

            # 3. Scroll to ensure visibility
            driver.execute_script("arguments[0].scrollIntoView({block: 'center', behavior: 'smooth'});",
                                  expanded_container)
            time.sleep(0.3)  # Small rendering pause

            # 4. Find all goal elements with retry logic
            goals = []
            for _ in range(3):  # Retry up to 3 times
                try:
                    goals = expanded_container.find_elements(By.XPATH, ".//br-fieldset-row")
                    if goals:
                        break
                    time.sleep(0.5)
                except StaleElementReferenceException:
                    continue

            if not goals:
                return []  # No goals found

            goal_data_list = []

            # 5. Process each goal with individual error handling
            for i, goal in enumerate(goals):
                try:
                    # Refresh reference to avoid staleness
                    goal = expanded_container.find_elements(By.XPATH, ".//br-fieldset-row")[i]

                    # Use JavaScript to get text content (more reliable for Angular)
                    legend_text = driver.execute_script("""
                    const legend = arguments[0].querySelector('legend');
                    return legend ? legend.innerText : '';
                    """, goal)

                    if not legend_text:
                        continue

                    # Parse the text content
                    lines = legend_text.split('\n')
                    if len(lines) < 2:
                        continue
                        
                    goal_data_list.append({
                        "meta": lines[4],
                        "unidade_medida": lines[5],
                        "quantidade": lines[6]
                    })

                except Exception as err:
                    print(f"⚠️ Erro ao extrair meta {i + 1}: {type(err).__name__} - {err}")
                    continue

            return goal_data_list

        except Exception as e:
            print(f"❌ Erro crítico ao extrair metas: {type(e).__name__} - {e}")
            return []

    try:
        if clicar_elemento(driver=driver, xpath=tabela_xpath):
            print(f"🔎📋 Tabela de executores localizada")
    except Exception as erro:
        print(f"⚠️ Tabela não encontrada: {type(erro).__name__} - {truncate_error(str(erro))}")
        return

    try:
        # Start inserting from this point
        insert_index = index
        # 📝 Save only the relevant columns to Excel
        colunas_para_salvar = ["Executor", "Objeto", "Meta", "Unidade de Medida", "Quantidade"]
        nova_linha_data = {
            "Executor": '',
            "Objeto": '',
            "Meta": '',
            "Unidade de Medida": '',
            "Quantidade": '',
        }

        try:
            # 2. Get first row only (no loop needed)
            rows = driver.find_elements(By.XPATH,tabela_xpath)
            if not rows:
                print("ℹ️ Nenhuma linha encontrada na tabela")
                return

            first_row = rows[0]  # Just process first row

            # Coleta a linha atual
            cells = first_row.find_elements(By.XPATH, ".//div[contains(@class, 'datatable-body-cell')]")

            # Get executor/objeto from first row
            try:
                executor = cells[1].text.strip() if len(cells) > 1 else ''
                objeto = cells[2].text.strip() if len(cells) > 2 else ''

                nova_linha_data.update({
                    "Executor": executor,
                    "Objeto": objeto,
                })
            except Exception as e:
                print(
                    f"⚠️ Dados não encontrados: {type(e).__name__} - {truncate_error(str(e))}")

            # Expand to get metas
            try:
                botao_expandir = cells[0].find_element(By.XPATH,
                                                       ".//button[contains(@class, 'br-button')]")
                botao_expandir.click()
            except Exception as e:
                print(f"❌ Não foi possível clicar para expandir a tabela")
                print(f"⚠️ Falha encontrada: {type(e).__name__} - {truncate_error(str(e))}")
            # Get metas
            try:
                metas = extract_metas_from_expanded_row()
            except Exception as e:
                metas = []
                print(f"❌ Erro ao coletar metas: {type(e).__name__} - {truncate_error(str(e),
                                                                        max_error_length=50)}")

            for meta in metas:
                nova_linha_data.update( {
                    "Meta": meta["meta"],
                    "Unidade de Medida": meta["unidade_medida"],
                    "Quantidade": meta["quantidade"],
                })
                try:
                    # Check if current row's columns are empty
                    current_row_empty = True
                    if insert_index < len(df):
                        current_row_empty = all(
                            pd.isna(df.at[insert_index, col])
                            or
                            df.at[insert_index, col] == "" for col in colunas_para_salvar
                        )
                    if current_row_empty and insert_index < len(df):
                        # Fill existing empty row
                        for col in colunas_para_salvar:
                            df.at[insert_index, col] = nova_linha_data[col]

                        print(f"ℹ️ Dados inseridos na linha existente {insert_index}")
                    else:
                        insert_index += 1

                        # Convert the new data to a DataFrame
                        new_row = {col: nova_linha_data.get(col, '') for col in df.columns}
                        new_row_df = pd.DataFrame([new_row])

                        # Split the original DataFrame
                        upper = df.iloc[:insert_index]
                        lower = df.iloc[insert_index:]

                        df = pd.concat([upper, new_row_df, lower], ignore_index=True)

                        print(f"ℹ️ Nova linha criada na posição {insert_index + 1}")
                except Exception as erro:
                    print(f"❌ Erro ao salvar listas no Excel: {type(erro).__name__} -"
                          f" {truncate_error(str(erro))}")

            df.to_excel(df_path, index=False)

        except Exception as erro:
            print(f"❌ Erro ao processar a linha.\n{type(erro).__name__} - {truncate_error(str(erro))}")
    except Exception as erro:
        print(f"❌ Erro ao coletar dados: {type(erro).__name__} - {truncate_error(str(erro))}")

# Lineariza o dicionário aninhado
def flatten_dict(d, parent_key='', sep='.'):
    """
    d → dicionário

    parent_key → prefixo

    sep → separador

    items → resultado

    k → chave

    v → valor

    new_key → caminho

    flatten_dict(...) → recursão
    """

    items = {}
    for k, v in d.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.update(flatten_dict(v, new_key, sep=sep)) # recursão
        else:
            items[new_key] = v
    return items

# Verifica se a checkbox está marcada
def esta_selecionado(driver, use_second_case=False):
    """
    Checks which radio button ("Sim" or "Não") is selected and returns it under the key 'declaracoes'.

    Args:
        driver: Selenium WebDriver instance.
        use_second_case (bool): If True, checks the second possible radio group (for dual cases).

    Returns:
        dict: {"declaracoes": "Sim" | "Não" | None} (None if none selected or error)
    """
    try:
        # Locate the radio group dynamically (adjust selector if needed)
        if not use_second_case:
            # Case 1: First radio group (default)
            radio_group = driver.find_element(By.XPATH,
                                              "//div[@class='field disabled']//div[contains(@class, 'br-radio')]")
        else:
            # Case 2: Second radio group (if exists)
            radio_group = driver.find_element(By.XPATH,
                                              "(//div[@class='field disabled']//div[contains(@class, 'br-radio')])[2]")

        # Find all radio inputs in the group
        radio_buttons = radio_group.find_elements(By.XPATH, ".//input[@type='radio']")

        # Check which one is selected
        for radio in radio_buttons:
            if radio.is_selected():
                label = radio.find_element(By.XPATH, "./following-sibling::label").text
                return {"declaracoes": label}  # "Sim" or "Não"

        # If none selected
        return {"declaracoes": None}

    except Exception as e:
        print(f"Error: {e}")
        return {"declaracoes": None}

# Sobe os dados para o arquivo excel
def atualiza_excel(df_path, df, index, plano_acao: dict,  col_range: list=None,
                   init_range: int=0, fin_range: int=0, second_init: bool=False):
    selected_keys = [
        "pagamentos.empenho",
        "pagamentos.valor",
        "pagamentos.ordem",
        "declaracoes.recursos_orcamento",
        "classificacao_orcamentaria",
        "declaracoes.nao_uso_pessoal",
        "prazo_de_execucao",
        "periodo_exec",
        "controle_social.conselhos",
        "controle_social.instancias"
    ]

    # Flatten the dictionary and filter only selected keys
    flat_dict = flatten_dict(plano_acao)
    filtered_data = {k: flat_dict.get(k, "") for k in selected_keys if flat_dict.get(k) not in [None, ""]}

    # Map to Excel columns
    column_headers = [
        "Empenho", "Valor", "Ordem do Pagamento", "Indicação Orçamento Beneficiario",
        "Classificação Orçamentária de Despesa","Declaração Recurso","Prazo Execução", "Período Execução"
                      ]
    if second_init:
        df = pd.read_excel(df_path, dtype=str)

    if col_range:
        columns_range = col_range
        selected_keys = list(flat_dict.keys())[init_range:fin_range]

        for col_idx, key in zip(columns_range, selected_keys):
            value = flat_dict.get(key, "")
            df[df.columns[col_idx]] = df[df.columns[col_idx]].astype('object')
            df.iat[index, col_idx] = value  # iat is used for fast scalar access

    else:
        # Update Excel cells
        for col_name, value in zip(column_headers, filtered_data.values()):
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name)
                df[df.columns[col_idx]] = df[df.columns[col_idx]].astype('object')
                df.iat[index, col_idx] = value
            else:
                print(f"⚠️ Column '{col_name}' not found in DataFrame")

    df.to_excel(df_path, index=False)

# Verify is the row has data in it
def has_actual_data(x):
    """Returns True if value is non-empty and non-NA"""
    if pd.isna(x):
        return False
    if isinstance(x, str):
        return x.strip() != ""
    return True  # Numbers, booleans, etc.

# Função principal
def main():
    driver = conectar_navegador_existente()

    planilha_final = (r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência "
                      r"Social\Teste001\Sofia\Emissão Parecer 2024 (732 PT) - Copia.xlsx")


    try:
        df = pd.read_excel(planilha_final, engine='openpyxl').astype(object)
        print(f"✅ Planilha lida com {len(df)} linhas.")

        # Clica no que menu de navegação
        clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/br-header/"
                                "header/div/div[2]/div/div[1]/button/span")
        # Clica em 'Plano de Ação'
        clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/"
                                "div/br-side-menu/nav/div[3]/a/span[2]")

        remover_backdrop(driver)

        # Clica no ícone para filtrar
        clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/"
                                "main/transferencia-especial-main/transferencia-plano-acao/transferencia-plano-acao-consulta/br-table/div/div/div/button/i")

        # Processa cada linha do DataFrame usando o índice
        index = 0
        #for index, row in df.iterrows():
        while index < len(df):
            print(len(df))

            row = df.iloc[index]
            plano_acao = {
                "beneficiario": {
                    "nome": "",  # Nome do beneficiário
                    "uf": ""  # Unidade Federativa (estado)
                },
                "dados_bancarios": {
                    "banco": "",  # Nome do banco
                    "agencia": "",  # Número da agência
                    "conta": "",  # Número da conta
                    "situacao": ""  # Situação atual da conta
                },
                "emenda": {
                    "numero": "",  # Número/identificação da emenda
                    "valor": ""  # Valor do investimento
                },
                "finalidade": ""
                ,
                "programacoes_orcamentarias": ""
                ,
                "pagamentos": {
                    "empenho": "",  # Número do empenho
                    "valor": "",  # Valor do empenho
                    "ordem": ""  # Número da ordem de pagamento
                },
                "declaracoes": {
                    "recursos_orcamento": False,  # Recursos no orçamento próprio?
                    "nao_uso_pessoal": False  # Não será usado para pessoal/dívida?
                },
                "execucao": {
                    "executor": "",  # Nome do executor
                    "objeto": "",  # Objeto do projeto
                },
                "metas": {
                    "descricao": "",  # Descrição da meta
                    "unidade": "",  # Unidade de medida
                    "quantidade": "",  # Quantidade prevista
                },
                "historico": {
                    "sistema": "",  # Histórico registrado no sistema
                    "concluido": ""  # Histórico após conclusão
                },
                "controle_social": {
                    "conselhos": "",  # Informações dos conselhos locais
                    "instancias": ""  # Instâncias de controle social
                },
                "periodo_exec ": "",
                "prazo_de_execucao": "",
                "classificacao_orcamentaria": ""
            }
            codigo = str(row["Código do Plano de Ação"])  # Garante que o código seja string

            # Verifica se o código já foi processado (coluna "Responsável" preenchida e diferente de erro)
            row_without_first_col = row.iloc[1:] # Excludes index 0 (first column)
            if row_without_first_col.apply(has_actual_data).any():
                print(f"⏭️ Pulando linha... {index} ")
                index += 1
                continue

            if pd.isna(df.at[index,"Código do Plano de Ação"]):
                index += 1
                continue

            print(f"\n{' Inicio do loop ':=^60}\n")
            print(f"Processando código: {codigo} (índice: {index})")

            try:
                # Insere o código no campo de filtro
                xpath_campo_filtro = ("/html/body/transferencia-especial-root/br-main-layout/"
                                      "div/div/div/main/transferencia-especial-main/transferencia-plano-acao/"
                                      "transferencia-plano-acao-consulta/br-table/div/br-fieldset/fieldset/"
                                      "div[2]/form/div[1]/div[2]/br-input/div/div[1]/input")
                if not inserir_texto(driver, xpath_campo_filtro, codigo):
                    raise Exception("Falha ao inserir o código no filtro")

                # Variáveis para o retry
                max_tentativas = 10  # Número máximo de tentativas
                tentativa = 1
                codigo_tabela_xpath = ("/html/body/transferencia-especial-root/br-main-layout/div/div/div"
                                       "/main/transferencia-especial-main/transferencia-plano-acao"
                                       "/transferencia-plano-acao-consulta/br-table/div/ngx-datatable/div"
                                       "/datatable-body/datatable-selection/datatable-scroller/datatable"
                                       "-row-wrapper/datatable-body-row/div[2]/datatable-body-cell[2]/div")
                codigo_correspondente = False



                while tentativa <= max_tentativas and not codigo_correspondente:
                    print(f"ℹ️ Tentativa {tentativa} de {max_tentativas} para filtrar o código {codigo}")

                    # Aplica o filtro
                    if not clicar_elemento(driver,
                                           "/html/body/transferencia-especial-root/br-main-layout/div/"
                                           "div/div/main/transferencia-especial-main/transferencia-plano-acao/"
                                           "transferencia-plano-acao-consulta/br-table/div/"
                                           "br-fieldset/fieldset/div[2]/form/div[5]/div/button[2]"):
                        raise Exception("Falha ao aplicar o filtro")

                    # Aguarda a tabela atualizar
                    WebDriverWait(driver, 7).until(
                        EC.presence_of_element_located((By.XPATH, codigo_tabela_xpath))
                    )
                    time.sleep(0.5)  # Pequeno delay adicional para garantir a atualização

                    # Verifica o código na tabela
                    codigo_tabela = obter_texto(driver, codigo_tabela_xpath)
                    if not codigo_tabela:
                        print(f"⚠️ Não foi possível obter o código na tabela na tentativa {tentativa}")
                    elif codigo_tabela.strip() == codigo.strip():
                        print(f"✅ Código verificado com sucesso: {codigo_tabela}")
                        codigo_correspondente = True
                    else:
                        print(f"⚠️ Código na tabela ({codigo_tabela}) não corresponde ao inserido ({codigo})"
                              f" na tentativa {tentativa}")
                        time.sleep(1)  # Espera antes de tentar novamente

                    tentativa += 1

                if not codigo_correspondente:
                    raise Exception(f"Falha ao filtrar o código {codigo} após {max_tentativas} tentativas."
                                    f" Último código na tabela: {codigo_tabela}")

                # Clica em "Detalhar" somente se o código estiver correto
                detalhar_xpath = ("/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/"
                                  "transferencia-especial-main/transferencia-plano-acao/"
                                  "transferencia-plano-acao-consulta/br-table/div/ngx-datatable/div/"
                                  "datatable-body/datatable-selection/datatable-scroller/"
                                  "datatable-row-wrapper/datatable-body-row/div[2]/datatable-body-cell[8]/"
                                  "div/div/button/i")

                if not clicar_elemento(driver, detalhar_xpath):
                    raise Exception("Falha ao clicar em 'Detalhar'")
                
                remover_backdrop(driver)

                # Coleta os dados da primeira pagina
                plano_acao = loop_primeira_pagina(driver=driver, plano_acao=plano_acao)
                domain = list(range(1,11))
                atualiza_excel(
                    df_path=planilha_final,  # Excel file path
                    df=df,  # DataFrame to modify
                    index=index,  # Row index to update (REQUIRED)
                    col_range=domain,
                    plano_acao=plano_acao,  # Dictionary containing data
                    init_range=0,  # First key to use from flattened dict
                    fin_range=10  # Last key to use (exclusive)
                )

                print(f"\n{' Inicio do loop da segunda página ':=^60}\n")

                # Coleta os dados da segunda pagina
                plano_acao = loop_segunda_pagina(driver=driver, plano_acao=plano_acao, index=index, df=df,
                                                 df_path=planilha_final)
                atualiza_excel(
                    df_path=planilha_final,  # Excel file path
                    df=df,  # DataFrame to modify
                    index=index,  # Row index to update (REQUIRED)
                    plano_acao=plano_acao,  # Dictionary containing data
                    second_init=True
                )

                print(f"\n{' Fim do loop ':=^60}\n")

                df = pd.read_excel(planilha_final, engine='openpyxl').astype(object)

                # Sobe para o topo da página
                driver.execute_script("window.scrollTo(0, 0);")
                # Volta para a seleção do plano de ação
                clicar_elemento(driver,
                                "/html/body/transferencia-especial-root/br-main-layout/div/div/div/main"
                                "/div[1]/br-breadcrumbs/div/ul/li[2]/a")
                # Clica em "Filtrar" para o próximo código
                clicar_elemento(driver,
                                "/html/body/transferencia-especial-root/br-main-layout/div/div/div/main"
                                "/transferencia-especial-main/transferencia-plano-acao/transferencia-plano"
                                "-acao-consulta/br-table/div/div/div/button/i")

                index += 1

            except Exception as erro:
                last_error = truncate_error(f"Main loop element intercepted: {str(erro)}")
                print(f"⚠️ {last_error}")

                try:
                    clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/"
                                            "div/div/div/main/transferencia-especial-main/"
                                            "transferencia-plano-acao/transferencia-plano-acao-consulta/"
                                            "br-table/div/div/div/button")
                except:
                    print("⚠️ Falha ao recuperar a tela de consulta. Reiniciando navegação...")
                    driver.get("about:blank")
                    clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/"
                                            "br-header/header/div/div[2]/div/div[1]/button/span")
                    clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/"
                                            "div/div/div/div/br-side-menu/nav/div[3]/a/span[2]")
                    remover_backdrop(driver)
                    clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/"
                                            "div/div/div/main/transferencia-especial-main/"
                                            "transferencia-plano-acao/transferencia-plano-acao-consulta/"
                                            "br-table/div/div/div/button/i")
                continue

        print("✅ Todos os dados foram coletados e salvos na planilha!")

    except Exception as erro:
        last_error = truncate_error(f"Element intercepted: {str(erro)}, {type(erro).__name__}")
        print(f"❌ {last_error}")

if __name__ == "__main__":
    main()