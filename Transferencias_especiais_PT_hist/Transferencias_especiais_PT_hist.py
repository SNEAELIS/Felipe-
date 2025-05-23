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
            options=chrome_options)
        driver.switch_to.window(driver.window_handles[0])
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
            print(f"✅ Clicked element (attempt {tentativa}): {elemento.text}...")

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

        # Método 1: Atributo value padrão
        try:
            WebDriverWait(driver, 2).until(
                lambda d: elemento.get_attribute("value") and elemento.get_attribute("value").strip() != "")
            valor = elemento.get_attribute("value")
        except TimeoutException:
            pass  # Vamos tentar outros métodos

        # Método 2: JavaScript para campos desabilitados
        if not valor:
            try:
                valor = driver.execute_script(
                    "return arguments[0].value || arguments[0].defaultValue;",
                    elemento)
            except Exception as js_err:
                print(f"⚠️ JS fallback falhou: {truncate_error(str(js_err))}")

        # 3. Tratamento do valor retornado
        if valor and str(valor).strip():
            valor = str(valor).strip()
            print(f"📦📄 Valor obtido: {valor[:50]}...")  # Trunca valores longos
            return valor
        else:
            print("ℹ️ Campo vazio ou sem valor válido")
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
        print("✅ Backdrop removido com sucesso!")
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
def extract_all_rows_text(driver, table_xpath, timeout=10, count: int = 0):
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

            # Print or process the extracted data
            print(f"Row {idx}: {row_data}")

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
def loop_segunda_pagina(driver, index, df_path):
    lista_caminhos = [
         # [0] > Aba Plano de Trabalho
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/'
        'transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/ul/li[3]/button/span',

        # [1] > Historico Sistema +  Historico Concluido
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-plano'
        '-trabalho/br-fieldset[2]/fieldset/div[2]/div/div/br-table/div/ngx-datatable',
    ]

    # Aba Plano de Trabalho
    clicar_elemento(driver, lista_caminhos[0])

    # Histórico section
    coletar_dados_hist(driver, lista_caminhos[1], index=index, df_path=df_path)

# Coleta os dados do histórico
def coletar_dados_hist(driver, tabela_xpath, index, df_path):
    df = pd.read_excel(df_path, dtype=str)
    try:
        # Wait for table to be present
        try:
            tabela = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, tabela_xpath))
            )
            print("ℹ️ Tabela de objetos localizada")
        except Exception as erro:
            print(f"⚠️ Tabela não encontrada: {erro}")
            return

        # Get all rows in the table body
        rows = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located(
                (By.XPATH, f"{tabela_xpath}//datatable-body-row")
            )
        )

        data = []
        colunas_para_salvar = ["Responsável_sys", "Data/Hora_sys", "Situação_sys",
                               "Responsável", "Data/Hora", "Situação"]

        if not rows:
            data = ["Não Encontrado"] * 6
        else:
            # Process each row to extract cell data
            table_data = []
            for row in rows:
                cells = row.find_elements(By.XPATH, ".//datatable-body-cell")
                row_data = []
                for cell in cells:
                    # Extract text content from each cell
                    cell_text = cell.text.strip()
                    row_data.append(cell_text)
                table_data.append(row_data)

            print(
                f"ℹ️ Found {len(table_data)} rows with {len(table_data[0]) if table_data else 0} columns each")

            # Search backwards for "Sistema" row
            for i in range(len(table_data) - 1, -1, -1):
                try:
                    if table_data[i][0] == "Sistema":
                        sys_data = table_data[i][:3]  # First 3 columns

                        # Look for "Concluído" in subsequent rows
                        conc_data = ["Não Encontrado"] * 3
                        for j in range(i + 1, len(table_data)):
                            if table_data[j][2] == "Concluído":
                                conc_data = table_data[j][:3]
                                break

                        data = sys_data + conc_data
                        break
                except Exception as erro:
                    print(f"❌ Erro ao processar linha {i}: {type(erro).__name__} - {str(erro)}")
                    continue

        # Save to DataFrame
        if data:
            for j, col_name in enumerate(colunas_para_salvar):
                df.at[index, col_name] = data[j]
            df.to_excel(df_path, index=False, engine='openpyxl')
            print(f"✅ Dados salvos em {df_path[-30:]}\n")
        else:
            print("⚠️ Nenhum dado relevante encontrado na tabela")

    except Exception as erro:
        print(f"❌ Erro ao coletar dados: {erro}")


# Função principal
def main():
    driver = conectar_navegador_existente()

    planilha_final = (r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência "
                      r"Social\Teste001\Sofia\PT SNEAELIS até dia 14_04_2025(SOFIA)"
                      r" - datas_hist - Copia.xlsx")

    try:
        df = pd.read_excel(planilha_final)
        print(f"✅ Planilha lida com {len(df)} linhas.")

        if df["Código do Plano de Ação"].duplicated().any():
            print("⚠️ Aviso: Há códigos duplicados na planilha. Removendo duplicatas...")
            df = df.drop_duplicates(subset=["Código do Plano de Ação"], keep="first")

        # Clica no que menu de navegação
        clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/br-header/"
                                "header/div/div[2]/div/div[1]/button/span")
        # Clica em 'Plano de Ação'
        clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/"
                                "div/br-side-menu/nav/div[3]/a/span[2]")

        remover_backdrop(driver)


        # Processa cada linha do DataFrame usando o índice
        for index, row in df.iterrows():
            # Verifica se o código já foi processado (coluna "Responsável" preenchida e diferente de erro)
            if pd.notna(df.at[index, "Situação"]):
                print(f"ℹ️ Linha {index} já tem situação de conclusão: {df.at[index, 'Situação-Conc']}")
                continue

            # Clica no ícone para filtrar
            clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/"
                                    "main/transferencia-especial-main/transferencia-plano-acao/transferencia-plano-acao-consulta/br-table/div/div/div/button/i")

            codigo = str(row["Código do Plano de Ação"])  # Garante que o código seja string

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

                # Navega até "Plano de Trabalho"
                remover_backdrop(driver)
                if not clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/"
                                               "div/div/div/main/transferencia-especial-main/"
                                               "transferencia-plano-acao/transferencia-cadastro/br-tab-set/"
                                               "div/nav/ul/li[3]/button"):
                    raise Exception("Falha ao navegar para 'Plano de Trabalho'")

                # Coleta os dados da segunda pagina
                loop_segunda_pagina(driver=driver, index=index, df_path=planilha_final)
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
        print(f'{type(erro).__name__}')


if __name__ == "__main__":
    main()


