from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (TimeoutException, NoSuchElementException,
                                        ElementClickInterceptedException, WebDriverException,
                                        StaleElementReferenceException, ElementNotInteractableException)
from selenium.webdriver.chrome.service import Service
from datetime import datetime
import pandas as pd
import time
import sys


# Função para conectar ao navegador já aberto
def conectar_navegador_existente():
    """
    Inicializa o Objeto Robo, configurando e iniciando o driver do Chrome.
    """
    try:
        # Configuração do registro
        # Inicia as opções do Chrome
        chrome_options = webdriver.ChromeOptions()
        # Endereço de depuração para conexão com o Chrome
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        # Inicializa o driver do Chrome com as opções e o gerenciador de drivers
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options)

        print("✅ Conectado ao navegador existente com sucesso.")

        return driver
    except WebDriverException as e:
        # Imprime mensagem de erro se a conexão falhar
        print(f"❌ Erro ao conectar ao navegador existente: {type(e).__name__}\nError: {str(e)[:200]}...")


# Trunca a mensagem de erro
def truncate_error(msg, max_error_length=100):
    """Truncates error message with ellipsis if too long"""
    return (msg[:max_error_length] + '...') if len(msg) > max_error_length else msg


# Função para clicar em um elemento com retry
def clicar_elemento(driver, xpath, retries=3, prt: bool = True):
    """Attempts to click an element with truncated error messages.

    Args:
        driver: WebDriver instance
        xpath: XPath of the element to click
        retries: Number of retry attempts
        prt: If the text should be printed or not
    """
    last_error = None

    for tentativa in range(1, retries + 1):
        try:
            remover_backdrop(driver)
            elemento = WebDriverWait(driver, 7).until(
                EC.element_to_be_clickable((By.XPATH, xpath)))
            if prt:
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
            last_error = truncate_error(f"Timeout (attempt {tentativa}): {str(e)[:30]}")
            print(f"⏱️ {last_error}")

        except NoSuchElementException as e:
            last_error = truncate_error(f"Element not found (attempt {tentativa}): {str(e)}")
            print(f"🔍 {last_error}")

        except Exception as e:
            last_error = truncate_error(f"Unexpected error (attempt {tentativa}): {str(e)}")
            print(f"❌ {last_error}")

        if tentativa < retries:
            time.sleep(tentativa)

    print(truncate_error(f"❌ Failed after {retries} attempts for {elemento.text}. Last error: {last_error}"))
    return False


# Função para inserir texto em um campo com retry
def inserir_texto(driver, xpath, texto, retries=3):
    """
        Insere texto em um campo de formulário identificado por XPath, com tentativas de repetição.

        Esta função tenta localizar um elemento em uma página web usando o XPath fornecido,
        remove eventuais elementos sobrepostos (backdrop), limpa o campo e insere o texto especificado.
        Em caso de falha, realiza novas tentativas até atingir o número máximo definido.

        Parameters:
            driver: Instância do WebDriver utilizada para controlar o navegador.
            xpath (str): XPath do campo onde o texto deve ser inserido.
            texto (str): Texto que será inserido no campo.
            retries (int, optional): Número máximo de tentativas em caso de erro. Padrão é 3.

        Returns:
            bool: Retorna True se o texto for inserido com sucesso, False caso contrário.
        """
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
            print(f"❌ Erro ao inserir texto no campo {xpath} (tentativa {tentativa + 1}): {erro[:50]}")
            time.sleep(1)  # Espera antes de tentar novamente
    print(f"❌ Falha após {retries} tentativas para inserir texto em {xpath}")
    return False


# Função para obter texto de um elemento
def obter_texto(driver, xpath):
    """
    Obtém o texto de um elemento da página identificado por XPath.

    Esta função aguarda a presença do elemento na página, remove eventuais sobreposições
    (como backdrop), extrai o texto e retorna o conteúdo tratado. Se o elemento não contiver
    texto visível ou ocorrer um erro, retorna None.

    Parameters:
        driver: Instância do WebDriver utilizada para controlar o navegador.
        xpath (str): XPath do elemento de onde o texto será extraído.

    Returns:
        str or None: Texto extraído do elemento, ou None em caso de erro ou se o texto estiver vazio.
    """
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
def get_radio_selection(driver):
    def get_angular_radio_state():
        """
        Robustly gets radio button state in Angular applications with:
        - Angular stability waiting
        - Multiple fallback methods
        - Comprehensive error handling
        - Automatic retries
        """

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

        def execute_with_fallbacks():
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
                        return str(result).lower()  # Normalize to string
                except:
                    continue

            return None

        def check_visual_state(element):
            """Fallback for visually determining state"""
            active_class = driver.execute_script("""
            return window.getComputedStyle(arguments[0]).getPropertyValue('active');
            """, element)
            return 'true' if active_class and 'active' in active_class else None

        max_attempts = 3
        for attempt in range(max_attempts):
            try:
                # 1. Wait for Angular to stabilize (critical for Angular apps)
                wait_for_angular()

                # 2. Wait for the specific element to be interactable
                radio_group = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, 'br-radio[formcontrolname="isRecursosIndicados"]'))
                )

                # 3. Try multiple detection methods with fallbacks
                value = execute_with_fallbacks()

                if value is not None:
                    return value

                time.sleep(1)  # Brief pause before retry

            except Exception as e:
                print(f"Attempt {attempt + 1} failed:", str(e))

        print("All attempts exhausted. Saving debug information...")
        return None

    try:
        radio_button = get_angular_radio_state()

        # Execute the script and get the result
        selected_value = radio_button

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
    """
           Obtém o valor de um campo desabilitado em uma página web usando várias estratégias.

           Esta função tenta localizar um campo (possivelmente desabilitado) usando seu XPath e obter
           seu valor por meio de diferentes abordagens: atributo `value`, execução de JavaScript ou
           leitura direta do texto visível. Se o campo não for encontrado ou não contiver valor, retorna
           um valor padrão ou None.

           Parameters:
               driver: Instância do WebDriver usada para controlar o navegador.
               xpath (str): XPath do campo de onde o valor deve ser extraído.
               call (int, optional): Número identificador da chamada, usado para fins de log/debug. Padrão é None.

           Returns:
               str or None: Valor extraído do campo, 'Campo Vazio' se não encontrado ou em caso de falhas,
               ou None se o campo estiver presente mas sem valor identificável.
    """
    try:
        remover_backdrop(driver)  # Remove o backdrop antes de obter valor
        try:
            elemento = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.XPATH, xpath)))
        except TimeoutException:
            print(f"⏱️ Timeout: Elemento não encontrado em 5s - {truncate_error(xpath, 500)}")
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
    """
    Remove elementos de sobreposição (backdrop) da página.

    Esta função executa um script JavaScript para remover todos os elementos com a classe
    'backdrop' do DOM. É útil para evitar bloqueios visuais ou funcionais em interações
    automatizadas com a interface da página.

    Parameters:
       driver: Instância do WebDriver usada para controlar o navegador.

    Returns:
        None
        """
    try:
        driver.execute_script("""
            var backdrops = document.getElementsByClassName('backdrop');
            while(backdrops.length > 0){
                backdrops[0].parentNode.removeChild(backdrops[0]);
            }
        """)
        # print("✅ Backdrop removido com sucesso!")
    except Exception as erro:
        print(f"❌ Erro ao remover backdrop: {erro}")


# Função que aplica a espera pelo elemento
def wait_for_element(driver, xpath: str, timeout: int = 5) -> bool:
    """Aguarda até que um elemento esteja presente na página."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath)))
        return True
    except Exception:
        return False


# Extrai os dados da tabela de forma concatenada
def extract_all_rows_text(driver, table_xpath, timeout=2.5):
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

        joined_row = " || ".join(row_data)

        print(f"📦📄 Valor obtido: {joined_row[:50]}...")  # Trunca valores longos

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
    """ Loop que pega dados da primeira aba """
    #  [0] Beneficiário // [1] UF // [2] Banco // [3] Agência // [4] Emenda Parlamentar
    #  [5] Valor de Investimento // [6] Valor de Custeio
    #  [8] Dados Complementares do Plano // [8] Finalidades
    lista_caminhos = [
        # Beneficiário [0][1]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[1]/fieldset/div[2]/div[1]/div[1]/br-input/div/div[1]/input',

        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[1]/fieldset/div[2]/div[1]/div[2]/br-input/div/div[1]/input',

        # Banco [2][3]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[1]/fieldset/div[2]/div[2]/div[1]/br-input/div/div[1]/input',

        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[1]/fieldset/div[2]/div[2]/div[2]/br-input/div/div[1]/input',

        # Dados emenda parlamentar [4] > [6]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[2]/fieldset/div[2]/div/div[1]/br-input/div/div[1]/input',

        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[2]/fieldset/div[2]/div/div[3]/br-input/div/div[1]/input',

        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[2]/fieldset/div[2]/div/div[2]/br-input/div/div[1]/input',

        # Dados Complementares do Plano [7]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[3]/fieldset/div[2]/div[1]/div/br-textarea/div/div[1]/div/textarea',
        
        # Finalidades [8]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-basicos/form/br-fieldset[3]/fieldset/div[2]',
    ]

    remover_backdrop(driver)
    WebDriverWait(driver, 7).until(
        EC.presence_of_element_located((By.XPATH,
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/plano-acao-info/div'
    )))

    plano_acao["beneficiario"]["nome"] = obter_valor_campo_desabilitado(driver, lista_caminhos[0])
    plano_acao["beneficiario"]["uf"] = obter_valor_campo_desabilitado(driver, lista_caminhos[1])

    plano_acao["dados_bancarios"]["banco"] = obter_valor_campo_desabilitado(driver, lista_caminhos[2])
    plano_acao["dados_bancarios"]["agencia"] = obter_valor_campo_desabilitado(driver, lista_caminhos[3])

    plano_acao["emenda"]["numero"] = obter_valor_campo_desabilitado(driver, lista_caminhos[4])
    plano_acao["emenda"]["valor"] = obter_valor_campo_desabilitado(driver, lista_caminhos[5])
    plano_acao["emenda"]["custeio"] = obter_valor_campo_desabilitado(driver, lista_caminhos[6])

    plano_acao["dados_comp_plano"] = obter_valor_campo_desabilitado(driver, lista_caminhos[7])

    plano_acao["finalidades"] = extract_all_rows_text(driver=driver, table_xpath=lista_caminhos[8])

    return plano_acao


# Navega pela primeira página e coleta os dados
def loop_segunda_pagina(driver, index, plano_acao: dict, df, df_path):
    """ Loop que pega dados da segunda aba em diante"""
    # [0] Aba Plano de Trabalho
    # [1] Declaracão Não Uso Pessoal
    # [2] Classificação Orçamentária de Despesa
    # [3] Prazo de Execução em meses
    # [4] Execucao Executor ; Objeto ; Metas

    lista_caminhos = [
        # [0]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/'
        'transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/ul/li[3]/button/span',

        # [1]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-plano'
        '-trabalho/form/div/div/br-fieldset[2]/fieldset/div[2]/div[1]/div/br-checkbox/div',

        # [2]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-plano'
        '-trabalho/form/div/div/br-fieldset[1]/fieldset/div[2]/div[3]/div/br-textarea/div/div['
        '1]/div/textarea',

        # [3]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-plano'
        '-trabalho/form/div/div/br-fieldset[2]/fieldset/div[2]/div[2]/div[1]/br-input/div/div[1]/input',

        # [4]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-plano'
        '-trabalho/form/div/div/br-fieldset[3]/fieldset/div[2]/div[2]/div[2]/br-table/div/ngx-datatable',
    ]

    try:
        # Aba Plano de Trabalho
        clicar_elemento(driver, lista_caminhos[0])
        remover_backdrop(driver)
        time.sleep(3)
        try:
            error_msg_path = ('/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/br'
                              '-alert-messages/div/div/div[2]/div')
            error_msg_elm = driver.find_element(By.XPATH, error_msg_path)

            if error_msg_elm:
                driver.refresh()
                time.sleep(3)
        except:
            pass

        # Declarações section
        plano_acao["declaracoes"]["recursos_orcamento"] = get_radio_selection(driver)
        plano_acao["declaracoes"]["nao_uso_pessoal"] = obter_valor_campo_desabilitado(driver,
                                                                                      lista_caminhos[1])
        if plano_acao["declaracoes"]["nao_uso_pessoal"] not in ['', 'Campo Vazio']:
            plano_acao["declaracoes"]["nao_uso_pessoal"] = 'Sim'
        else:
            plano_acao["declaracoes"]["nao_uso_pessoal"] = 'Não'

        # Classificação Orçamentária de Despesa
        plano_acao["classificacao_orcamentaria"] = obter_valor_campo_desabilitado(driver, lista_caminhos[2])

        # Prazo de Execução em meses
        plano_acao["prazo_de_execucao"] = obter_valor_campo_desabilitado(driver, lista_caminhos[3])

        # Execução and Metas section
        coletar_dados_listas(driver, lista_caminhos[4], index=index, df=df, df_path=df_path)

        return plano_acao

    except Exception as erro:
        print(f"❌ Erro inesperado: {type(erro).__name__} - {truncate_error(str(erro))}")
        return None


# Função para coletar executores e metas
def coletar_dados_listas(driver, tabela_xpath, index, df, df_path):
    """
       Coleta dados de uma tabela expandida e insere as informações em um DataFrame, salvando no formato Excel.

       A função localiza uma tabela via XPath, coleta dados das células (incluindo informações expandidas)
       e insere ou atualiza essas informações no DataFrame passado como parâmetro. Se a linha a ser atualizada
       já contiver dados, ela será preenchida. Caso contrário, uma nova linha será criada no DataFrame.

       Parameters:
           driver: Instância do WebDriver para controle da página e coleta de dados.
           tabela_xpath (str): XPath que localiza a tabela na página.
           index (int): Índice da linha no DataFrame onde os dados devem ser inseridos ou atualizados.
           df (pd.DataFrame): DataFrame que contém as linhas de dados a serem atualizadas.
           df_path (str): Caminho para o arquivo Excel onde o DataFrame será salvo.

       Side Effects:
           - Interage com a página web, clicando para expandir a tabela.
           - Atualiza o DataFrame com os dados extraídos da tabela.
           - Salva o DataFrame modificado em um arquivo Excel.

       Returns:
           None: A função salva os dados no arquivo Excel e não retorna nenhum valor.

       Exceptions:
           - Exceções podem ser lançadas se houver problemas ao localizar a tabela ou ao processar os dados.
           - Caso ocorram falhas ao clicar para expandir ou ao extrair dados das metas, a função tentará continuar o processo.
       """
    print(f"\n🔧 Processando Lista Executores — índice {index}\n")

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

    # Extracts all metas from the currently expanded executor row
    def extract_metas_from_expanded_row():
        """
        Extracts all metas from the currently expanded executor row.
        Returns a list of dictionaries: each with meta, unidade_medida, quantidade.
        """
        try:
            remover_backdrop(driver)
            _wait_for_angular_stable()
            metas_path = ("/html/body/transferencia-especial-root/br-main-layout/div/div/div/main"
                          "/transferencia-especial-main/transferencia-plano-acao/transferencia-cadastro/br"
                          "-tab-set/div/nav/transferencia-plano-trabalho-executor/br-tab-set/div/nav/br-tab"
                          "/div/form/br-fieldset[1]/fieldset/div[2]/div[3]/div[2]/br-table/div/ngx-datatable")

            # Find all goal elements with retry logic
            goals = None
            for _ in range(10):  # Retry up to 3 times
                try:
                    goals = driver.find_elements(By.XPATH, f"{metas_path}//datatable-body-row")
                    if goals:
                        break
                    time.sleep(1)
                except StaleElementReferenceException:
                    continue

            # Get element location
            element_y = goals[0].location['y']
            navbar_height = 100  # Adjust based on your UI
            # Scroll to element position minus navbar offset
            driver.execute_script(f"window.scrollTo(0, {element_y - navbar_height});")

            # Process each goal with individual error handling
            goal_data_list = []
            for goal in goals:
                goal_data_list_partial = []
                try:
                    goal_cells_data = goal.find_elements(By.XPATH, ".//datatable-body-cell")
                    # Extract text from inner div or span
                    for goal_data in goal_cells_data[1:]:
                        # Some cells might have <div>, some might have <span>
                        goal_texts = goal_data.find_elements(By.XPATH, './div | ./span')

                        for goal_col in goal_texts:
                            goal_data_list_partial.append(goal_col.text.strip())

                    # Append first row od data
                    goal_data_list.append({
                        "meta": goal_data_list_partial[0],
                        "descrição": goal_data_list_partial[1],
                        "unidade_medida": goal_data_list_partial[2],
                        "quantidade": goal_data_list_partial[3],
                        "meses_previstos": goal_data_list_partial[4]
                    })

                    expand_button = goal_cells_data[0].find_element(By.XPATH,
                                                                    ".//button[contains(@class, 'br-button')]")
                    expand_button.click()
                    try:
                        expanded_container = WebDriverWait(driver, 10).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "div.expanded-container")))
                        # Skip headers
                        details = expanded_container.find_elements(By.CSS_SELECTOR,
                                                                   "div.row:not(:first-child)")
                        # Get detailed "meta" data
                        for detail in details:
                            cols = detail.find_elements(By.CSS_SELECTOR, "div[class^='col-sm-']")
                            detail_data = [col.text.strip() for col in cols]
                            goal_data_list.append({
                                "categoria": detail_data[0],
                                "emenda_Especial": detail_data[1],
                                "recurso_Próprio": detail_data[2],
                                "rendimento_aplicação": detail_data[3],
                                "doações": detail_data[4],
                            })
                    except Exception as err:
                        print(f"⚠️ Erro ao extrair detalhes: {type(err).__name__}")
                        exc_type2, exc_obj2, exc_tb2 = sys.exc_info()
                        line_number2 = exc_tb2.tb_lineno
                        print(f"⚠️ Erro na linha {line_number2}: {type(err).__name__} - {err}")
                        continue

                    # Closes detailed "meta"
                    expand_button.click()

                except Exception as err:
                    print(f"⚠️ Erro ao extrair meta: {type(err).__name__}")
                    exc_type3, exc_obj3, exc_tb3 = sys.exc_info()
                    line_number3 = exc_tb3.tb_lineno
                    print(f"⚠️ Erro na linha {line_number3}: {type(err).__name__} - {err}")
                    continue
            print(f"⛏️🎲 Metas Lista Executores Extraídos")
            return goal_data_list

        except Exception as e:
            print(f"❌ Erro crítico ao extrair metas: {type(e).__name__} - {e}")
            return []

    try:
        if clicar_elemento(driver=driver, xpath=tabela_xpath, prt=False):
            pass
    except Exception as erro:
        print(f"⚠️ Tabela não encontrada: {type(erro).__name__} - {truncate_error(str(erro))}")
        return

    # Start inserting from this point
    insert_index = index
    # 📝 Save only the relevant columns to Excel
    colunas_para_salvar = [
        "Executor",
        "Objeto",
        "Meta",
        "Descrição",
        "Unidade de Medida",
        "Quantidade",
        "Meses Previstos",
        "Categoria",
        "Emenda Especial",
        "Recurso Próprio",
        "Rendimento de Aplicação",
        "Doações",
        "Email",

    ]
    # Get data before expanding "metas" this is done to avoid duplicating these two information
    executor_dict = {
        "Executor": '',
        "Objeto": '',
    }
    # Get data from "metas"
    nova_linha_data = {
        "Meta": '',
        "Descrição": '',
        "Unidade_medida": '',
        "Quantidade": '',
        "Meses_previstos": '',
        "Categoria": '',
        "Emenda_Especial": '',
        "Recurso_Próprio": '',
        "Rendimento_aplicação": '',
        "Doações": '',
        "Email": '',
    }
    emails_data = ''
    try:
        try:
            # Find table to count how many rows it has to iterate over
            rows = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located(
                (By.XPATH, f"{tabela_xpath}//datatable-body-row")))

            if not rows:
                print("ℹ️ Nenhuma linha encontrada na tabela")
                return
            # Iterate the first table to access the "metas"
            for row_index, row in enumerate(rows):
                remover_backdrop(driver)
                # Refind elements, to avoid stale element exception
                rows = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located(
                    (By.XPATH, f"{tabela_xpath}//datatable-body-row")))
                # Coleta a linha atual
                cells = rows[row_index].find_elements(By.XPATH, ".//div[contains(@class, "
                                                                "'datatable-body-cell')]")
                # Get executor/objeto from first row
                try:
                    executor = cells[1].text.strip() if len(cells) > 1 else ''
                    objeto = cells[2].text.strip() if len(cells) > 2 else ''

                    executor_dict.update({
                        "Executor": executor,
                        "Objeto": objeto,
                    })
                except Exception as e:
                    print(
                        f"⚠️ Dados não encontrados: {type(e).__name__} - {truncate_error(str(e))}")

                # Expand to get executor detailed data
                try:
                    remover_backdrop(driver)
                    botao_expandir = cells[5].find_element(By.XPATH,
                                                           ".//button[contains(@class, 'br-button')]")
                    botao_expandir.click()
                except Exception as e:
                    print(f"❌ Não foi possível clicar para expandir a tabela")
                    print(f"⚠️ Falha encontrada: {type(e).__name__} - {truncate_error(str(e))}")
                try:
                    # Get e-mails Lista de Conselhos locais ou instâncias de controle social
                    emails_path = (r'/html/body/transferencia-especial-root/br-main-layout/div/div/div/main'
                                   r'/transferencia-especial-main/transferencia-plano-acao/transferencia'
                                   r'-cadastro/br-tab-set/div/nav/transferencia-plano-trabalho-executor/br'
                                   r'-tab-set/div/nav/br-tab/div/form/br-fieldset[3]/fieldset/div[2]/div['
                                   r'1]/div[2]/br-table/div/ngx-datatable')

                    emails = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located(
                        (By.XPATH, f"{emails_path}//datatable-body-row")))

                    if len(emails) > 2:
                        emails_data = extract_all_rows_text(driver, emails_path)[1:]
                    else:
                        emails_data = obter_texto(driver, emails_path).split('\n')[-1]

                except Exception as e:
                    print(f"⚠️ Dados não encontrados: {type(e).__name__} - {truncate_error(str(e))}")

                # Get metas
                try:
                    metas = extract_metas_from_expanded_row()
                except Exception as e:
                    metas = []
                    print(f"❌ Erro ao coletar metas: {type(e).__name__} - {truncate_error(str(e),
                                                                                          max_error_length=50)}")
                go_back_button_path = ('/html/body/transferencia-especial-root/br-main-layout/div/div/div'
                                       '/main/transferencia-especial-main/transferencia-plano-acao'
                                       '/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano'
                                       '-trabalho-executor/br-tab-set/div/nav/br-tab/div/div/button')
                # retorna para a página de plano de trabalho
                driver.find_element(By.XPATH, go_back_button_path).click()

                nova_linha_data.update({"Email": emails_data})
                # Insert the initial nova_linha_data (email/social) to the DataFrame
                # Just-in-time merge for DataFrame insertion
                row_to_insert = {**executor_dict, **nova_linha_data}
                # Insert the merged row, filling empty cells where needed
                current_row_empty = True
                if insert_index < len(df):
                    current_row_empty = all(
                        pd.isna(df.at[insert_index, col]) or df.at[insert_index, col] == "" for col in
                        colunas_para_salvar
                    )
                if current_row_empty and insert_index < len(df):
                    # Fill existing empty row
                    for col in colunas_para_salvar:
                        df.at[insert_index, col] = row_to_insert.get(col, "")
                else:
                    insert_index += 1
                    # Convert the new data to a DataFrame
                    new_row = {col: row_to_insert.get(col, "") for col in df.columns}
                    new_row_df = pd.DataFrame([new_row])
                    # Split the original DataFrame
                    upper = df.iloc[:insert_index]
                    lower = df.iloc[insert_index:]
                    df = pd.concat([upper, new_row_df, lower], ignore_index=True)

                # Reset social data so it's not duplicated
                nova_linha_data.update({"Email": ''})

                for i in range(0, len(metas), 3):
                    meta_main = metas[i]
                    meta_custeio = metas[i + 1]
                    meta_investimento = metas[i + 2]

                    # Reset values
                    nova_linha_data.update({
                        "Meta":  "",
                        "Descrição":  "",
                        "Unidade de Medida":  "",
                        "Quantidade": "",
                        "Meses Previstos":  ""
                    })

                    # Now, you can safely build your nova_linha_data and financial_rows like this:
                    nova_linha_data.update({
                        "Meta": meta_main.get("meta", ""),
                        "Descrição": meta_main.get("descrição", ""),
                        "Unidade de Medida": meta_main.get("unidade_medida", ""),
                        "Quantidade": meta_main.get("quantidade", ""),
                        "Meses Previstos": meta_main.get("meses_previstos", "")
                    })
                    financial_rows = [
                        {
                            "Categoria": meta_custeio.get("categoria", ""),
                            "Emenda Especial": meta_custeio.get("emenda_Especial", ""),
                            "Recurso Próprio": meta_custeio.get("recurso_Próprio", ""),
                            "Rendimento de Aplicação": meta_custeio.get("rendimento_aplicação", ""),
                            "Doações": meta_custeio.get("doações", ""),
                        },
                        {
                            "Categoria": meta_investimento.get("categoria", ""),
                            "Emenda Especial": meta_investimento.get("emenda_Especial", ""),
                            "Recurso Próprio": meta_investimento.get("recurso_Próprio", ""),
                            "Rendimento de Aplicação": meta_investimento.get("rendimento_aplicação", ""),
                            "Doações": meta_investimento.get("doações", ""),
                        }
                    ]

                    try:
                        # Add financial values first
                        for f in financial_rows:
                            colunas_para_salvar_fin = colunas_para_salvar[7:12]
                            # Check if current row's columns are empty
                            current_row_empty = True
                            if insert_index < len(df):
                                current_row_empty = all(
                                    pd.isna(df.at[insert_index, col])
                                    or
                                    df.at[insert_index, col] == "" for col in colunas_para_salvar_fin
                                )
                            if current_row_empty and insert_index < len(df):
                                # Fill existing empty row
                                for col in colunas_para_salvar_fin:
                                    df.at[insert_index, col] = f.get(col, "")

                            else:
                                insert_index += 1

                                # Convert the new data to a DataFrame
                                new_row = {col: f.get(col, '') for col in df.columns}
                                new_row_df = pd.DataFrame([new_row])

                                # Split the original DataFrame
                                upper = df.iloc[:insert_index]
                                lower = df.iloc[insert_index:]

                                df = pd.concat([upper, new_row_df, lower], ignore_index=True)

                        # Add 'metas' data
                        current_row_empty = True
                        if insert_index < len(df):
                            current_row_empty = all(
                                pd.isna(df.at[insert_index, col])
                                or
                                df.at[insert_index, col] == "" for col in (colunas_para_salvar[:7] +
                                                                           colunas_para_salvar[13:]))
                        # Select columns
                        colunas_para_salvar_sel = colunas_para_salvar[:7] + colunas_para_salvar[13:]

                        if current_row_empty and insert_index < len(df):
                            # Fill existing empty row
                            for col in colunas_para_salvar_sel:
                                df.at[insert_index, col] = nova_linha_data.get(col, "")

                        else:
                            insert_index += 1

                            # Convert the new data to a DataFrame
                            new_row = {col: nova_linha_data.get(col, '') for col in df.columns}
                            new_row_df = pd.DataFrame([new_row])

                            # Split the original DataFrame
                            upper = df.iloc[:insert_index]
                            lower = df.iloc[insert_index:]

                            df = pd.concat([upper, new_row_df, lower], ignore_index=True)

                    except Exception as erro:
                        print(f"❌ Erro ao salvar listas no Excel: {type(erro).__name__} -"
                              f" {truncate_error(str(erro))}")
            # Send data to excel
            df.to_excel(df_path, index=False)
        except Exception as erro:
            exc_type, exc_obj, exc_tb = sys.exc_info()
            line_number = exc_tb.tb_lineno
            print(f"❌ Erro ao extrair as linhas. Erro na linha: {line_number}\n{type(erro).__name__} -"
                  f" {truncate_error(str(erro))}")
    except Exception as erro:
        print(f"❌ Erro ao coletar dados: {type(erro).__name__} - {truncate_error(str(erro))}")


# Lineariza o dicionário aninhado
def flatten_dict(dictionary, parent_key='', sep='.'):
    """
    parent_key → prefixo
    sep → separador
    items → resultado
    k → chave
    v → valor
    new_key → caminho
    flatten_dict(...) → recursão
    """

    items = {}
    for k, v in dictionary.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.update(flatten_dict(dictionary=v, parent_key=new_key, sep=sep))  # recursão
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
def atualiza_excel(df_path, df, index, plano_acao: dict, col_range: list = None,
                   init_range: int = 0, fin_range: int = 0, second_init: bool = False):
    """
        Atualiza os dados de um arquivo Excel com informações extraídas de um dicionário.

        A função processa um dicionário contendo dados (plano de ação), filtra as chaves relevantes,
        e atualiza as células de um DataFrame com os valores extraídos. O DataFrame é então salvo de volta
        no arquivo Excel especificado.

        Parâmetros:
            df_path (str): Caminho para o arquivo Excel a ser atualizado.
            df (pd.DataFrame): DataFrame contendo os dados a serem atualizados no arquivo Excel.
            index (int): Índice da linha no DataFrame onde os dados devem ser inseridos ou atualizados.
            plano_acao (dict): Dicionário contendo os dados a serem extraídos e atualizados.
            col_range (list, opcional): Lista com os índices das colunas a serem atualizadas. Se `None`,
                                         a função atualizará as colunas padrão.
            init_range (int, opcional): Índice inicial para a seleção das chaves do dicionário a serem usadas,
                                         caso `col_range` seja fornecido.
            fin_range (int, opcional): Índice final para a seleção das chaves do dicionário a serem usadas,
                                       caso `col_range` seja fornecido.
            second_init (bool, opcional): Se `True`, recarrega o DataFrame do arquivo Excel antes de atualizar.

        Retorna:
            None: A função modifica diretamente o DataFrame e salva no arquivo Excel.

        Exceções:
            - Se alguma chave do dicionário não for encontrada no DataFrame, um aviso será impresso.
            - Se o arquivo Excel não puder ser acessado ou salvo, uma exceção será gerada.

        Observações:
            - O dicionário `plano_acao` é achatado antes de ser usado.
            - As chaves a serem usadas são previamente selecionadas e mapeadas para colunas no Excel.
            - Caso `col_range` seja fornecido, as colunas selecionadas serão atualizadas com os dados correspondentes.

        """

    selected_keys = [
        "declaracoes.recursos_orcamento",
        "classificacao_orcamentaria",
        "declaracoes.nao_uso_pessoal",
        "prazo_de_execucao",
    ]

    # Flatten the dictionary and filter only selected keys
    flat_dict = flatten_dict(plano_acao)
    filtered_data = {k: flat_dict.get(k, "") for k in selected_keys if flat_dict.get(k) not in [None, ""]}

    # Map to Excel columns
    column_headers = [
        "Indicação Orçamento Beneficiario","Classificação Orçamentária de Despesa",
        "Declaração Recurso", "Prazo Execução"
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


# Align the rows of the "Metas" column and save a new file with the current date appended to the name
def align_meta_and_fields_with_next_custeio(output_filename=None):
    # Read the Excel file
    filename = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\Teste001\Sofia\Emendas pix 2025\Emendas pix 2025_2_ciclo - Copia.xlsx')

    df = pd.read_excel(filename, dtype=str)

    # The fields to move together with "Meta"
    columns_to_move = [
        'Meta',
        'Descrição',
        'Unidade de Medida',
        'Quantidade',
        'Meses Previstos'
    ]

    # Check columns exist
    for col in columns_to_move + ['Categoria']:
        if col not in df.columns:
            raise ValueError(f"The DataFrame must contain the column: {col}")

    # Buffer for the values to move
    buffer = None
    last_custeio_idx = None

    for idx, row in df.iterrows():
        if row['Categoria'] == 'Custeio':
            last_custeio_idx = idx
        if row['Categoria'] == 'Investimento' and pd.notna(row['Meta']) and str(row['Meta']).strip():
            # Buffer the entire set of fields
            buffer = {col: row[col] for col in columns_to_move}
            # Clear these fields from the Investimento row
            for col in columns_to_move:
                df.at[idx, col] = ""
                df.at[last_custeio_idx, col] = buffer[col]
            buffer = None  # Clear the buffer

    # Output file name
    if output_filename is None:
        if filename.lower().endswith('.xlsx'):
            output_filename = f'{filename[:-13]}_{datetime.now():%d_%m_%Y}.xlsx'

    # Save the aligned DataFrame
    df.to_excel(output_filename, index=False)
    print(f"Aligned file saved as: {output_filename}")

# Função principal
def main():
    driver = conectar_navegador_existente()

    planilha_final = (r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência "
                      r"Social\Teste001\Sofia\Emendas pix 2025\Emendas pix 2025_2_ciclo - Copia.xlsx")

    df = pd.read_excel(planilha_final, engine='openpyxl').astype(object)
    print(f"✅ Planilha lida com {len(df)} linhas.")

    try:
        # Volta para a pagina inicial
        clicar_elemento(driver, '/html/body/transferencia-especial-root/br-main-layout/br-header/'
                                'header/div/div[2]/div/div[2]/div[1]/a')
        time.sleep(2)
        # Clica no que menu de navegação
        clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/br-header/"
                                "header/div/div[2]/div/div[1]/button/span")
        # Clica em 'Plano de Ação'
        clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/"
                                "div/br-side-menu/nav/div[3]/a/span[2]")

        remover_backdrop(driver)

        # Clica no ícone para filtrar
        clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/"
                                "main/transferencia-especial-main/transferencia-plano-acao/transferencia"
                                "-plano-acao-consulta/br-table/div/div/div/button/i")

        # Processa cada linha do DataFrame usando o índice
        index = 0
        # for index, row in df.iterrows():
        while index < len(df):
            print(f'Procentagem completa: {(index/len(df) * 100)}%'.center(50))

            row = df.iloc[index]
            plano_acao = {
                "beneficiario": {
                    "nome": "",  # Nome do beneficiário
                    "uf": ""  # Unidade Federativa (estado)
                },
                "dados_bancarios": {
                    "banco": "",  # Nome do banco
                    "agencia": "",  # Número da agência
                },
                "emenda": {
                    "numero": "",  # Número/identificação da emenda
                    "valor": "",  ## Valor de investimento
                    "custeio": ""  # Valor de custeio
                },
                "dados_comp_plano": ""
                ,
                "finalidades": ""
                ,
                "declaracoes": {
                    "recursos_orcamento": False,  # Recursos no orçamento próprio?
                    "nao_uso_pessoal": False  # Não será usado para despesa de pessoal/ serviço da dívida?
                },
                "execucao": {
                    "executor": "",  # Nome do executor
                    "objeto": "",  # Objeto do projeto
                },
                "metas": {
                    "descricao": "",  # Descrição da meta
                    "unidade": "",  # Unidade de medida
                    "quantidade": "",  # Quantidade prevista
                    "meses_previstos": "", # Meses previstos
                },
                "prazo_de_execucao": "",
                "classificacao_orcamentaria": "",
            }

            codigo = str(row["Código do Plano de Ação"])  # Garante que o código seja string

            # Verifica se o código já foi processado (coluna "Responsável" preenchida e diferente de erro)
            row_without_first_col = row.iloc[1:]  # Excludes index 0 (first column)
            if row_without_first_col.apply(has_actual_data).any():
                print(f"⏭️ Pulando linha... {index} ")
                index += 1
                continue

            if pd.isna(df.at[index, "Código do Plano de Ação"]):
                index += 1
                continue

            print(f"\n{' Inicio do loop ':=^60}\n")
            print(f"Processando código: {codigo} (índice: {index})\n")

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
                                       "-row-wrapper/datatable-body-row/div[2]/datatable-body-cell[2]"
                                       "/div/span")

                # Verifica se o código na tabela filtrada é igual ao código no campo de filtro
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
                detalhar_xpath = ("/html/body/transferencia-especial-root/br-main-layout/div/div/div/main"
                                  "/transferencia-especial-main/transferencia-plano-acao/transferencia"
                                  "-plano-acao-consulta/br-table/div/ngx-datatable/div/datatable-body"
                                  "/datatable-selection/datatable-scroller/datatable-row-wrapper/datatable"
                                  "-body-row/div[2]/datatable-body-cell[9]/div/div/button")

                if not clicar_elemento(driver, detalhar_xpath):
                    raise Exception("Falha ao clicar em 'Detalhar'")

                remover_backdrop(driver)

                print(f"\n{' Inicio do loop da primeira página ':=^60}\n")
                # Coleta os dados da primeira pagina
                plano_acao = loop_primeira_pagina(driver=driver, plano_acao=plano_acao)
                domain = list(range(1, 10))
                atualiza_excel(
                    df_path=planilha_final,  # Excel file path
                    df=df,  # DataFrame to modify
                    index=index,  # Row index to update (REQUIRED)
                    col_range=domain,
                    plano_acao=plano_acao,  # Dictionary containing data
                    init_range=0,  # First key to use from flattened dict
                    fin_range=9  # Last key to use (exclusive)
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
        align_meta_and_fields_with_next_custeio()
        driver.quit()
    except Exception as erro:
        last_error = truncate_error(f"Element intercepted: {str(erro)}, {type(erro).__name__}")
        print(f"❌ {last_error}")


if __name__ == "__main__":
    main()