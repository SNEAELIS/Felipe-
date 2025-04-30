from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (TimeoutException, NoSuchElementException,
                                        ElementClickInterceptedException, WebDriverException,
                                        StaleElementReferenceException)
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
    wait = WebDriverWait(driver, 10)

    try:
        # Find the parent container of both radio options
        container = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))

        # Check both options by examining their parent divs
        sim_div = container.find_element(By.XPATH, ".//div[contains(@class, 'br-radio')][1]")
        nao_div = container.find_element(By.XPATH, ".//div[contains(@class, 'br-radio')][2]")

        # Check which one has the 'selected' class or other visual indicator
        if "selected" in sim_div.get_attribute("class").lower():
            print(f"🔘 Opção selecionada: Sim")
            return "Sim"
        elif "selected" in nao_div.get_attribute("class").lower():
            print(f"🔘 Opção selecionada: Não")
            return "Não"

        # Alternative approach: check for checked attribute on hidden inputs
        sim_input = sim_div.find_element(By.XPATH, ".//input")
        if sim_input.get_attribute("checked"):
            print(f"🔘 Opção selecionada: Sim")
            return "Sim"

        nao_input = nao_div.find_element(By.XPATH, ".//input")
        if nao_input.get_attribute("checked"):
            print(f"🔘 Opção selecionada: Não")
            return "Não"

        # Fallback to JavaScript check
        if driver.execute_script("return arguments[0].checked;", sim_input):
            print(f"🔘 Opção selecionada: Sim")
            return "Sim"
        if driver.execute_script("return arguments[0].checked;", nao_input):
            print(f"🔘 Opção selecionada: Não")
            return "Não"

        # Default if nothing is selected
        print(f"🔘 Opção selecionada: Não")
        return "Não"

    except Exception as e:
        print(f"⚠️ Error checking radio buttons: {str(e)}")
        print(f"🔘 Opção selecionada: Não")
        return "Não"


# Função para obter o valor de um campo desabilitado
def obter_valor_campo_desabilitado(driver, xpath):
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
                print(f"❌ Erro inesperado ao acessar elemento.text: {type(e).__name__} - {e}")

        # Tratamento do valor retornado
        if valor and str(valor).strip():
            valor = str(valor).strip()
            print(f"📦📄 Valor obtido: {valor[:50]}...")  # Trunca valores longos
            return valor
        else:
            print(f"ℹ️ Campo:{elemento.text} vazio ou sem valor válido")
            return "Campo Vazio"
    except StaleElementReferenceException:
        print(f"👻 Elemento tornou-se obsoleto - {truncate_error(xpath)}")
        return None
    except Exception as erro:
        print(f"❌ Erro inesperado: {type(erro).__name__} - {truncate_error(str(erro))}")
        return None

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
        print(f"✅ Extracted joined data from table:\n{joined_row}")


        return joined_row

    except TimeoutException:
        print(f"❌ Timeout: No rows appeared within {timeout} seconds")
        return None
    except Exception as e:
        print(f"❌ Extraction failed: {str(e)[:100]}...")
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

    plano_acao["programacoes_orcamentarias"] = obter_valor_campo_desabilitado(driver, lista_caminhos[9])

    return plano_acao

# Navega pela primeira página e coleta os dados
def loop_segunda_pagina(driver,index, plano_acao: dict, df, df_path):
    # [0] Aba Dados Orçamentários // [1] Pagamentos Empenho // [2] Pagamentos Valor //
    # [3] Pagamentos Ordem // [4] Aba Plano de Trabalho // [5] Declaracoes Recursos Orcamento //
    # [6] Declaracoes Nao Uso Pessoal // [7] Execucao Executor // [8] Execucao Objeto //
    # [9] Metas Descricao // [10] Metas Quantidade // [11] Metas Unidade //
    # [12] Metas Meses // [13] Historico Sistema // [14] Historico Concluido //
    # [15] Controle Social Última notificação // [16] Controle Social Resp // [17] Período de Execução
    # [18]
    lista_caminhos = [
        # [0]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/'
        'transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/ul/li[2]/button/span',

        # [1]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/'
        'transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/'
        'transferencia-plano-acao-dados-orcamentarios/br-table[2]/div/ngx-datatable/div/datatable-body/'
        'datatable-selection/datatable-scroller/datatable-row-wrapper/datatable-body-row/div[2]/'
        'datatable-body-cell[1]/div/span',

        # [2]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-orcamentarios/br-table[2]/div/ngx-datatable/div/datatable-body/datatable-selection/'
        'datatable-scroller/datatable-row-wrapper/datatable-body-row/div[2]/datatable-body-cell[4]/div/span',

        # [3]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-dados'
        '-orcamentarios/br-table[2]/div/ngx-datatable/div/datatable-body/datatable-selection/datatable-'
        'scroller/datatable-row-wrapper/datatable-body-row/div[2]/datatable-body-cell[6]/div/button',

        # [4]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/'
        'transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/ul/li[3]/button/span',

        # [5]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-plano'
        '-trabalho/form/div/div/br-fieldset[1]/fieldset/div[2]/div[2]/div/br-radio/div',

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
        plano_acao["pagamentos"]["empenho"] = obter_valor_campo_desabilitado(driver, lista_caminhos[1])
        plano_acao["pagamentos"]["valor"] = obter_valor_campo_desabilitado(driver, lista_caminhos[2])
        plano_acao["pagamentos"]["ordem"] = obter_valor_campo_desabilitado(driver, lista_caminhos[3])

        # Aba Plano de Trabalho
        clicar_elemento(driver, lista_caminhos[4])

        # Declarações section
        plano_acao["declaracoes"]["recursos_orcamento"] = get_radio_selection(driver, lista_caminhos[5])
        time.sleep(0.5)
        plano_acao["declaracoes"]["nao_uso_pessoal"] = obter_valor_campo_desabilitado(driver, lista_caminhos[6])

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
    df = pd.read_excel(df_path, dtype=str)
    colunas_para_salvar = ["Responsável", "Data", "Situação"]  # Only 3 columns now

    try:
        # Wait for table to be present
        tabela = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, tabela_xpath))
        )
        print("🔎📋 Tabela de histórico localizada")

        # Get all rows in the table body
        rows = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located(
                (By.XPATH, f"{tabela_xpath}//datatable-body-row")
            )
        )

        if not rows:
            sys_data = ["Não Encontrado"] * 3
            conc_data = ["Não Encontrado"] * 3
        else:
            # Process each row to extract cell data
            table_data = []
            for row in rows:
                cells = row.find_elements(By.XPATH, ".//datatable-body-cell")
                row_data = [cell.text.strip() for cell in cells]
                table_data.append(row_data)

            print(f"ℹ️ Found {len(table_data)} rows with {len(table_data[0]) if table_data else 0} columns")

            # Initialize data variables
            sys_data = ["Não Encontrado"] * 3
            conc_data = ["Não Encontrado"] * 3

            # Search backwards for "Sistema" row
            for i in range(len(table_data) - 1, -1, -1):
                if table_data[i][0] == "Sistema":
                    sys_data = table_data[i][:3]
                    break

            # Search for "Concluído" row (after Sistema if possible)
            start_idx = i + 1 if 'i' in locals() else 0
            for j in range(start_idx, len(table_data)):
                if table_data[j][2] == "Concluído":
                    conc_data = table_data[j][:3]
                    break

        # Determine where to save the data
        current_row_empty = all(pd.isna(df.at[index, col]) or df.at[index, col] == ''
                                for col in colunas_para_salvar)

        if current_row_empty:
            # Save sys_data in current row
            for j, col_name in enumerate(colunas_para_salvar):
                df.at[index, col_name] = sys_data[j]

            # Check if next row is empty for conc_data
            next_index = index + 1
            if next_index >= len(df):
                df.loc[next_index] = [None] * len(df.columns)

            next_row_empty = all(pd.isna(df.at[next_index, col]) or df.at[next_index, col] == ''
                                 for col in colunas_para_salvar)

            if next_row_empty:
                # Save conc_data in next row
                for j, col_name in enumerate(colunas_para_salvar):
                    df.at[next_index, col_name] = conc_data[j]
            else:
                # Find next empty row after next_index
                new_index = next_index + 1
                while new_index < len(df) and not all(
                        pd.isna(df.at[new_index, col]) or df.at[new_index, col] == ''
                        for col in colunas_para_salvar):
                    new_index += 1

                if new_index >= len(df):
                    df.loc[new_index] = [None] * len(df.columns)

                # Save conc_data in new row
                for j, col_name in enumerate(colunas_para_salvar):
                    df.at[new_index, col_name] = conc_data[j]

        else:
            # Find first empty row after index
            new_index = index + 1
            while new_index < len(df) and not all(
                    pd.isna(df.at[new_index, col]) or df.at[new_index, col] == ''
                    for col in colunas_para_salvar):
                new_index += 1

            if new_index >= len(df):
                df.loc[new_index] = [None] * len(df.columns)

            # Save sys_data in first empty row
            for j, col_name in enumerate(colunas_para_salvar):
                df.at[new_index, col_name] = sys_data[j]

            # Find next empty row after new_index for conc_data
            conc_index = new_index + 1
            if conc_index >= len(df):
                df.loc[conc_index] = [None] * len(df.columns)

            # Save conc_data
            for j, col_name in enumerate(colunas_para_salvar):
                df.at[conc_index, col_name] = conc_data[j]

        # Save the DataFrame
        df.to_excel(df_path, index=False, engine='openpyxl')
        print(f"✅ Dados salvos no arquivo {df_path[-30:]}\n")

    except Exception as erro:
        print(f"❌ Erro ao coletar dados: {erro}")

# Opera o botão de expandir div da lista, caso necessário
def expand_executor_section_if_collapsed(driver):
    """Ensures the 'Dados do Executor' section is expanded."""
    wait = WebDriverWait(driver, 10)

    # Step 1: Locate the fieldset-header div that contains the legend 'Dados do Executor'
    fieldset_header = wait.until(EC.presence_of_element_located((
        By.XPATH,
        "//div[contains(@class, 'fieldset-header') and .//legend[text()='Dados do Executor']]"
    )))

    # Step 2: Locate the <i> toggle icon within that specific fieldset-header
    toggle_icon = fieldset_header.find_element(By.XPATH, ".//i[contains(@class, 'fa-angle')]")

    # Step 3: Check if it's collapsed
    icon_class = toggle_icon.get_attribute("class")
    if "fa-angle-down" in icon_class:
        toggle_icon.click()
        print("🔽 List was collapsed — toggled to expand.")
    else:
        print("👌 List already expanded — no action taken.")


# Função para coletar executores e metas
def coletar_dados_listas(driver, tabela_xpath, index, df, df_path):
    # Extracts all metas from the currently expanded executor row
    def extract_metas_from_expanded_row():
        """
        Extracts all metas from the currently expanded executor row.
        Returns a list of dictionaries: each with meta, unidade_medida, quantidade, meses_previstos.
        """
        wait = WebDriverWait(driver, 10)

        # Wait for the expanded section to become visible
        expanded_container = wait.until(EC.presence_of_element_located((
            By.XPATH, "//div[contains(@class, 'datatable-row-detail') and contains(@style, 'block')]"
                      "//div[contains(@class, 'expanded-container')]"
        )))

        # Find all metas (each one is a <br-fieldset-row>)
        goals = expanded_container.find_elements(By.XPATH, ".//br-fieldset-row")
        goal_data_list = []

        for goal in goals:
            try:
                # The data is stored in two rows inside the <legend>
                legend = goal.find_element(By.XPATH, ".//legend")
                lines = legend.find_elements(By.XPATH, ".//div[@class='row']")

                if len(rows) < 2:
                    continue  # Skip if structure isn't valid

                data_cells = lines[1].find_elements(By.XPATH, ".//div")

                # Validate we have the expected number of columns
                if len(data_cells) < 5:
                    continue

                goal_data_list.append({
                    "meta": data_cells[1].text.strip(),
                    "unidade_medida": data_cells[2].text.strip(),
                    "quantidade": data_cells[3].text.strip(),
                    "meses_previstos": data_cells[4].text.strip()
                })

            except Exception as err:
                print(f"⚠️ Erro ao extrair meta: {type(err).__name__} - {err}")
                continue

        return goal_data_list

    try:
       if clicar_elemento(driver=driver, xpath=tabela_xpath):
            print(f"🔎📋 Tabela de executores localizada")

       expand_executor_section_if_collapsed(driver=driver)
    except Exception as erro:
        print(f"⚠️ Tabela não encontrada: {type(erro).__name__} - {truncate_error(str(erro))}")
        return

    try:
        insert_index = index  # Start inserting from this point
        row_num = 1

        # 📝 Save only the relevant columns to Excel
        colunas_para_salvar = ["Executor", "Objeto", "Meta", "Unidade de Medida", "Quantidade",
                               "Meses Previstos"]

        while True:
            executor = ''
            objeto = ''
            nova_linha_data = {
                "Executor": '',
                "Objeto": '',
                "Meta": '',
                "Unidade de Medida": '',
                "Quantidade": '',
                "Meses Previstos": ''
            }
            try:
                # Get all visible rows again on each iteration
                rows = driver.find_elements(By.XPATH,
                       "//datatable-row-wrapper[contains(@class, 'datatable-row-wrapper')]")

                if row_num > len(rows):
                    print(f"✅ Todas as {len(rows)} linhas processadas. Encerrando.")
                    break

                print(f"📦 Processando linha {row_num}...")
                row = rows[row_num]

                # Scroll into view for stability (especially with virtual tables)
                driver.execute_script("arguments[0].scrollIntoView(true);", row)
                print(row.text)

                # Coleta a linha atual
                cells = row.find_elements(By.XPATH, ".//datatable-body-cell")


                if len(cells) < 3:
                    print(f"⚠️ Linha {row_num} não tem células suficientes. Pulando.")
                    row_num += 1
                    continue

                executor = cells[1].text.strip()
                objeto = cells[2].text.strip()

                nova_linha_data = {
                    "Executor": executor,
                    "Objeto": objeto,
                }

                if not objeto and not executor:
                    print(f"⚠️ Nenhum texto encontrado na linha {row_num}. Pulando...")
                    row_num += 1
                    continue
                time.sleep(10)
                try:
                    botao_expandir = cells[0].find_element(By.XPATH,
                                                           ".//button[contains(@class, 'br-button')]")
                    botao_expandir.click()
                except:
                    print(f"❌ Não foi possível clicar para expandir a linha {row_num}")
                    row_num += 1
                    continue

                # Aguarda as metas/etapas carregarem
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(
                        (By.XPATH,
                         "//div[contains(@class, 'datatable-row-detail') and contains(@style, 'block')]")
                    )
                )

                metas = extract_metas_from_expanded_row(driver)

                for meta in metas:
                    nova_linha_data = {
                        "Meta": meta["meta"],
                        "Unidade de Medida": meta["unidade_medida"],
                        "Quantidade": meta["quantidade"],
                        "Meses Previstos": meta["meses_previstos"]
                    }

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
                            df.to_excel(df_path, index=False)
                        print(f"ℹ️ Dados inseridos na linha existente {insert_index}")
                    else:
                        # Create new row
                        upper = df.iloc[:insert_index + 1]
                        lower = df.iloc[insert_index + 1:]
                        nova_linha_df = pd.DataFrame([nova_linha_data])
                        df = pd.concat([upper, nova_linha_df, lower], ignore_index=True)
                        df.to_excel(df_path, index=False)
                        print(f"ℹ️ Nova linha criada na posição {insert_index + 1}")
                        insert_index += 1

                row_num += 1
                time.sleep(0.5)  # Pequeno delay entre iterações para evitar sobrecarga

                try:
                    df[colunas_para_salvar].to_excel(df_path, index=False, engine='openpyxl')
                    print(f"✅ Dados salvos em {df_path}")
                except Exception as erro:
                    print(f"❌ Erro ao salvar Excel: {type(erro).__name__} - {truncate_error(str(erro))}")
            except Exception as erro:
                print(f"❌ Erro ao processar objeto {row_num}: {type(erro).__name__} -"
                      f" {truncate_error(str(erro))}")
                break
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
def atualiza_excel(df_add, df, index, plano_acao: dict,  col_range: list=None,
                   init_range: int=0, fin_range: int=0 ):
    selected_keys = [
        "pagamentos.empenho",
        "pagamentos.valor",
        "pagamentos.ordem",
        "declaracoes.recursos_orcamento",
        "declaracoes.nao_uso_pessoal",
        "controle_social.conselhos",
        "controle_social.instancias",
        "periodo_exec"
    ]


    # Flatten the dictionary and filter only selected keys
    flat_dict = flatten_dict(plano_acao)
    filtered_data = {k: flat_dict.get(k, "") for k in selected_keys if flat_dict.get(k) not in [None, ""]}

    # Map to Excel columns (adjust column_headers to match your Excel)
    column_headers = ["Empenho", "Valor", " Ordem do Pagamento", "Indicação Orçamento Beneficiario",
                      "Declaração Recurso", "Prazo Execução", "Período Execução"]
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

    df.to_excel(df_add, index=False)



# Função principal
def main():
    driver = conectar_navegador_existente()
    caminho_planilha = (r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e "
                        r"Assistência Social\Teste001\PT SNEAELIS até dia 14_04_2025(SOFIA).xlsx")

    planilha_final = (r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e "
                        r"Assistência Social\Teste001\PT SNEAELIS até dia 14_04_2025(SOFIA) - Copia.xlsx")

    try:
        df = pd.read_excel(caminho_planilha)
        print(f"✅ Planilha lida com {len(df)} linhas.")

        if df["Código do Plano de Ação "].duplicated().any():
            print("⚠️ Aviso: Há códigos duplicados na planilha. Removendo duplicatas...")
            df = df.drop_duplicates(subset=["Código do Plano de Ação "], keep="first")

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
        for index, row in df.iterrows():
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
                    "meses": ""  # Meses previstos
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

            codigo = str(row["Código do Plano de Ação "])  # Garante que o código seja string
            # Verifica se o código já foi processado (coluna "Responsável" preenchida e diferente de erro)
            if pd.notna(df.at[index, "Dados dos Conselhos locais ou instâncias de controle social"]):
                print(f"ℹ️ Linha {index} já tem situação de conclusão: {df.at[index,
                'Dados dos Conselhos locais ou instâncias de controle social']}")
                continue

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

                '''
                # Coleta os dados da primeira pagina
                plano_acao = loop_primeira_pagina(driver=driver, plano_acao=plano_acao)
                domain = list(range(1,11))
                atualiza_excel(
                    df_add=planilha_final,  # Excel file path
                    df=df,  # DataFrame to modify
                    index=index,  # Row index to update (REQUIRED)
                    col_range=domain,
                    plano_acao=plano_acao,  # Dictionary containing data
                    init_range=0,  # First key to use from flattened dict
                    fin_range=10  # Last key to use (exclusive)
                )

                print(f"\n{' Fim do loop da primeira página ':=^60}\n")

                '''
                # Coleta os dados da segunda pagina
                plano_acao = loop_segunda_pagina(driver=driver, plano_acao=plano_acao, index=index, df=df,
                                                 df_path=planilha_final)
                atualiza_excel(
                    df_add=planilha_final,  # Excel file path
                    df=df,  # DataFrame to modify
                    index=index,  # Row index to update (REQUIRED)
                    plano_acao=plano_acao,  # Dictionary containing data
                )

                # Sobe para o topo da página
                driver.execute_script("window.scrollTo(0, 0);")
                # Clica em "Filtrar" para o próximo código
                clicar_elemento(driver,
                                "/html/body/transferencia-especial-root/br-main-layout/div/div/div/main"
                                "/div[1]/br-breadcrumbs/div/ul/li[2]/a")

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
        last_error = truncate_error(f"Element intercepted: {str(erro)}")
        print(f"❌ {last_error}")

if __name__ == "__main__":
    main()