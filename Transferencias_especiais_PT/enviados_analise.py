import time
import win32com.client as win32
import os
import sys

import pandas as pd

from datetime import datetime

from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (TimeoutException, NoSuchElementException,
                                        ElementClickInterceptedException, WebDriverException,
                                        StaleElementReferenceException)
from selenium.webdriver.chrome.service import Service
from Transferencias_especiais_PT import skip_chrome_tab_search

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
        # Endereço de depuração para conexão com o Chromea
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9224")
        # Inicializa o driver do Chrome com as opções e o gerenciador de drivers
        driver = webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options)

        skip_chrome_tab_search(driver=driver)


        print("✅ Conectado ao navegador existente com sucesso.")

        return driver
    except WebDriverException as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred at line: {exc_tb.tb_lineno}")
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
            remover_backdrop(driver, False)
            elemento = WebDriverWait(driver, 7).until(
                EC.element_to_be_clickable((By.XPATH, xpath)))
            print(f"✅ Clicked element (attempt {tentativa}): {elemento.text}...")

            if elemento.is_displayed():
                elemento.click()
                return True
        except StaleElementReferenceException:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print("🔄 Stale element detected! Attempting recovery...")  # Debug print
            # Wait for the old element to go stale, then re-find
            WebDriverWait(driver, 10).until(EC.staleness_of(elemento))
            new_element = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, xpath)))
            new_element.click()

        except ElementClickInterceptedException as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            last_error = truncate_error(f"Element intercepted on click (attempt {tentativa}): {str(e)}")
            print(f"⚠️ {last_error}")

        except TimeoutException as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            last_error = truncate_error(f"Timeout (attempt {tentativa}): {str(e)}")
            print(f"⏱️ {last_error}")

        except NoSuchElementException as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            last_error = truncate_error(f"Element not found (attempt {tentativa}): {str(e)}")
            print(f"🔍 {last_error}")

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
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
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
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
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred at line: {exc_tb.tb_lineno}")
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
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"⏱️ Timeout: Elemento não encontrado em 5s - {truncate_error(xpath)}")
            return "Campo Vazio"
        except NoSuchElementException:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
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
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            pass  # Vamos tentar outros métodos

        # Método 2: JavaScript para campos desabilitados
        if not valor:
            try:
                valor = driver.execute_script(
                    "return arguments[0].value || arguments[0].defaultValue;",
                    elemento)
            except Exception as js_err:
                exc_type, exc_value, exc_tb = sys.exc_info()
                print(f"Error occurred at line: {exc_tb.tb_lineno}")
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
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred at line: {exc_tb.tb_lineno}")
        print(f"👻 Elemento tornou-se obsoleto - {truncate_error(xpath)}")
        return None
    except Exception as erro:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred at line: {exc_tb.tb_lineno}")
        print(f"❌ Erro inesperado: {type(erro).__name__} - {truncate_error(str(erro))}")
        return None


# Função para remover o backdrop
def remover_backdrop(driver, show_msg=True):
    try:
        driver.execute_script("""
            var backdrops = document.getElementsByClassName('backdrop');
            while(backdrops.length > 0){
                backdrops[0].parentNode.removeChild(backdrops[0]);
            }
        """)
        if show_msg:
            print("✅ Backdrop removido com sucesso!")
    except Exception as erro:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred at line: {exc_tb.tb_lineno}")
        print(f"❌ Erro ao remover backdrop: {erro}")


# Função que aplica a espera pelo elemento
def wait_for_element(driver, xpath: str, timeout: int = 10) -> bool:
    """Aguarda até que um elemento esteja presente na página."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath)))
        return True
    except Exception:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred at line: {exc_tb.tb_lineno}")
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
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred at line: {exc_tb.tb_lineno}")
        print(f"❌ Timeout: No rows appeared within {timeout} seconds")
        return None
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred at line: {exc_tb.tb_lineno}")
        print(f"❌ Extraction failed: {str(e)[:100]}...")
        return None


# Navega pela primeira página e coleta os dados
def loop_segunda_pagina(driver, index, df_path, sheet_name, df) -> bool:
    lista_caminhos = [
         # [0] > Aba Plano de Trabalho
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main/'
        'transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/ul/li[3]/button/span',

        # [1] > Historico Sistema +  Historico Concluido
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-plano-acao-plano'
        '-trabalho/br-fieldset[2]/fieldset/div[2]',
    ]

    # Aba Plano de Trabalho
    clicar_elemento(driver, lista_caminhos[0])

    # Histórico section
    if coletar_dados_hist(driver, lista_caminhos[1], index=index, df_path=df_path, sheet_name=sheet_name,
                          df=df):
        return True

    return False


# Coleta os dados do histórico
def coletar_dados_hist(driver, tabela_xpath, index, df_path, sheet_name, df) -> bool:
    try:
        # Get all rows in the table body
        rows = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located(
                (By.XPATH, f"{tabela_xpath}//datatable-body-row")
            )
        )

        data = []
        colunas_para_salvar = ["Responsável", "Data/Hora", "Situação"]

        if not rows:
            data = ["Não Encontrado"] * 3
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

            try:
                if table_data[0][2] == "Enviado para análise":
                    data = table_data[0][:3]  # First 3 cells

            except Exception as erro:
                exc_type, exc_value, exc_tb = sys.exc_info()
                print(f"Error occurred at line: {exc_tb.tb_lineno}")
                print(f"❌ Erro ao processar a tabela: {type(erro).__name__}\n{str(erro)[:50]}")

        # Save to DataFrame
        if data:
            for j, col_name in enumerate(colunas_para_salvar):
                df.at[index, col_name] = data[j]
            df.to_excel(df_path,sheet_name=sheet_name, index=False, engine='openpyxl')
            print(f"✅ Dados salvos em {df_path[-30:]}\n")

            return True
        else:
            print("⚠️ Nenhum dado relevante encontrado na tabela")
            return False

    except Exception as erro:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred at line: {exc_tb.tb_lineno}")
        print(f"❌ Erro ao coletar dados: {type(erro).__name__}")


def reset_browser(driver):
    try:
        # Clica no que menu de navegação
        clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/br-header/header/div"
                                "/div[2]/div/div[1]/button")
        # Clica em 'Plano de Ação'
        clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/div/br"
                                "-side-menu/nav/div[3]/a")
        print("🚀 Iniciando processo: acessando consultas de planos de ação.")
        remover_backdrop(driver)
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred at line: {exc_tb.tb_lineno}")
        print(f'❌ Falha ao resetar navegador.\nFalha{str(e)[:80]}')


# Função principal
def main(xlsx_path):
    driver = conectar_navegador_existente()

    planilha_final = xlsx_path

    try:
        all_sheets = pd.read_excel(planilha_final, sheet_name=None)

        for sheet_name, df in all_sheets.items():
            print(f"\n{'=' * 50}")
            print(f"📊 Processando planilha: {sheet_name}\n N° de linhas:{len(df)}.")
            print(f"{'=' * 50}")
            try:
                drop_rows = []
                for i, r in df.iterrows():
                    if len(df.columns) > 2:
                        drop_rows.append(i)
                df.drop(drop_rows, inplace=True)
                df.to_excel(planilha_final,sheet_name=sheet_name, index=False, engine='openpyxl')
                print(f"✅ Nº de dados removidos da planilha '{sheet_name}': {len(drop_rows)}")
            except Exception as e:
                exc_type, exc_value, exc_tb = sys.exc_info()
                print(f"Error occurred at line: {exc_tb.tb_lineno}")
                print(f'{type(e).__name__}\n{str(e)[:80]}')

            print(f"✅ Planilha {sheet_name} lida com {len(df)} linhas.")

            if df["Código do Plano de Ação"].duplicated().any():
                print("⚠️ Aviso: Há códigos duplicados na planilha. Removendo duplicatas...")
                df = df.drop_duplicates(subset=["Código do Plano de Ação"], keep="first")

            reset_browser(driver)
            # Processa cada linha do DataFrame usando o índice
            num_new_data = 0
            for index, row in df.iterrows():
                codigo = str(row["Código do Plano de Ação"])  # Garante que o código seja string
                print(f"\n⚙️ Processando código: {codigo} (índice: {index})\n")

                # Clica no ícone para filtrar
                clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/"
                                        "main/transferencia-especial-main/transferencia-plano-acao/transferencia-plano-acao-consulta/br-table/div/div/div/button/i")

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
                                               "div/div/main/transferencia-especial-main/"
                                               "transferencia-plano-acao/transferencia-plano-acao-consulta/"
                                               "br-table/div/br-fieldset/fieldset/div[2]/form/div[5]/"
                                               "div/button[2]"):

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

                    # Navega até "Plano de Trabalho"
                    remover_backdrop(driver)
                    if not clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/"
                                                   "div/div/div/main/transferencia-especial-main/"
                                                   "transferencia-plano-acao/transferencia-cadastro/br-tab-set/"
                                                   "div/nav/ul/li[3]/button"):
                        raise Exception("Falha ao navegar para 'Plano de Trabalho'")

                    # Coleta os dados da segunda pagina
                    if loop_segunda_pagina(driver=driver,
                                           index=index,
                                           df_path=planilha_final,
                                           sheet_name=sheet_name,
                                           df=df
                    ):
                        num_new_data += 1
                    # Sobe para o topo da página
                    driver.execute_script("window.scrollTo(0, 0);")
                    # Clica em "Filtrar" para o próximo código
                    clicar_elemento(driver,
                                    "/html/body/transferencia-especial-root/br-main-layout/div/div/div/main"
                                    "/div[1]/br-breadcrumbs/div/ul/li[2]/a")

                except Exception as erro:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(f"Error occurred at line: {exc_tb.tb_lineno}")
                    last_error = truncate_error(f"Main loop element intercepted: {str(erro)}")
                    print(f"⚠️ {last_error}")

                    try:
                        clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/"
                                                "div/div/div/main/transferencia-especial-main/"
                                                "transferencia-plano-acao/transferencia-plano-acao-consulta/"
                                                "br-table/div/div/div/button")
                    except:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(f"Error occurred at line: {exc_tb.tb_lineno}")
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

            print(f"✅ Todos os dados foram coletados e salvos na planilha!\n {num_new_data} "
                  f"novos dados adicionados")

    except Exception as erro:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred at line: {exc_tb.tb_lineno}")
        last_error = truncate_error(f"Element intercepted: {str(erro)}")
        print(f"❌ {last_error}")
        print(f'{type(erro).__name__}')



def send_emails_from_excel(excel_path):
    """
    Prepares emails from Excel data (Columns A, B, C) and opens them in Outlook for review.

    Args:
        excel_path (str): Path to Excel file.
        attachment_path (str, optional): Attachment file path.
        sender (str, optional): Email sender address.
    """

    def prepare_outlook_email(email:str, subject:str, html_body:str, attachment_paths:str):
        """Creates and displays an Outlook email (without sending)."""
        try:
            if not all(isinstance(x, str) for x in [email, subject, html_body]):
                raise TypeError("Recipient, subject, and body must be strings")

            outlook = win32.Dispatch('outlook.application')
            time.sleep(2)
            e_mail = outlook.CreateItem(0)

            e_mail.To = email
            e_mail.Subject = subject
            e_mail.HTMLBody = html_body

            if attachment_paths:
                # Convert single string to list for uniform handling
                if isinstance(attachment_paths, str):
                    attachment_paths = [attachment_paths]
                for path in attachment_paths:
                    if os.path.exists(path):
                        e_mail.Attachments.Add(path)
                        print(f"📎 Attached: {os.path.basename(path)}")
                    else:
                        print(f"⚠️ Attachment not found: {path}")

            e_mail.Display()  # Opens email for review (instead of .Send())
            print(f"📧 Email prepared for: {email}")
            return True
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Failed to prepare email for {email}: \n{e}")
            return False

    def generate_email_body(extra_data_list):
        """Generates HTML body from Excel data only."""
        today = datetime.now().strftime("%d/%m/%Y")
        subject = f"Atualização das diligencias. Data {today}"
        html_body = f"""<!DOCTYPE html>
        <html>
        <head>
            <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        </head>
        <body style="font-family: 'Courier New', monospace; margin: 0; padding: 0;">
        <div style="padding: 10px; background-color: #f0f0f0; border-bottom: 1px solid #ccc;
         font-weight: bold;">
            Arquivo: {os.path.basename(excel_path)}
        </div>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
                <td>"""
        for data in extra_data_list:
           ''' html_body += (
                f"<p style='white-space: pre ; font-family:Courier New, monospace>Código do Plano de "
                f"Ação: {data['A']}      "
                f"Responsável:"
                f" {data['B']}    "
                f"Data/Hora: {data['C']}    Situação: {data['D']}</p>"
                "<br><br></p>"  # 2 empty lines between entries
            )'''
           html_body += f"""
                       <div style="white-space: pre; margin-bottom: 20px;">
           Código do Plano de Ação:    {data['A']}
           Responsável:               {data['B']}
           Data/Hora:                 {data['C']}
           Situação:                  {data['D']}
                       </div>
                       """

           # Close all tags
           html_body += """</td>
               </tr>
            </table>
           </body>
           </html>"""
        return html_body, subject

    def read_excel_data():
        """Reads Excel and extracts Columns A, B, C if C has data."""
        try:
            df = pd.read_excel(excel_path, dtype=str)
            email = "maria.dourado@esporte.gov.br"
            extra_data_list = []  # Stores {A, B, C} dicts

            for _, row in df.iterrows():
                # Check if Column C (index 2) has data
                if pd.notna(row.iloc[3]):
                    extra_data = {
                        'A': row.iloc[0],  # Column A
                        'B': row.iloc[1],  # Column B
                        'C': row.iloc[2],  # Column C
                        'D': row.iloc[3]  # Column D
                    }
                    extra_data_list.append(extra_data)

            print(f'Número de planos de ação no email: {len(extra_data_list)}')
            if extra_data_list:
                return email, extra_data_list
            else:
                return [], []

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Failed to read Excel file: \n{e}")
            return [], []

    # Read data
    email, extra_data_list = read_excel_data()

    try:
        if email or extra_data_list:
        # Prepare emails
            html_body, subject = generate_email_body(extra_data_list)  # Pass single entry
            prepare_outlook_email(email=email,subject=subject,html_body=html_body,attachment_paths=excel_path)
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred at line: {exc_tb.tb_lineno}")
        print(f"❌ No data to prepare email for {type(e).__name__}: \n{str(e)[:100]}")


if __name__ == "__main__":
    xlsx_paths = [
        r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Transferencias_Especiais\Conferência robô 30abr MANHA.xlsx"
    ]

    for idx, path in enumerate(xlsx_paths):
        main(path)
        send_emails_from_excel(path)