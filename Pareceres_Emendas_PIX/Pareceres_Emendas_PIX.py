import sys

from fontTools.subset.svg import xpath
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (TimeoutException, NoSuchElementException,
                                        ElementClickInterceptedException, WebDriverException,
                                        StaleElementReferenceException)
from selenium.webdriver.chrome.service import Service
from functools import wraps
import pandas as pd
import time
import os
import numpy as np


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

        print("✅ Conectado ao navegador existente com sucesso.")

        reset_web_page(driver=driver)

        return driver
    except WebDriverException as e:
        # Imprime mensagem de erro se a conexão falhar
        print(f"❌ Erro ao conectar ao navegador existente: {e}")


# Reset the page to start the iteration process
def reset_web_page(driver):
    try:
        home_page_icon_path = ('/html/body/transferencia-especial-root/br-main-layout/br-header/header/div/div['
                               '2]/div/div[2]/div[1]/a')
        clicar_elemento(driver=driver, xpath=home_page_icon_path)
        print('De volta a pagina inicial !')
        return True
    except Exception as e:
        print(f"Error:{type(e).__name__}. Error_txt: {str(e)[:100]}")
        return False

# Trunca a mensagem de erro
def truncate_error(msg, max_error_length=100):
    """Truncates error message with ellipsis if too long"""
    return (msg[:max_error_length] + '...') if len(msg) > max_error_length else msg


# Função para clicar em um elemento com retry
def clicar_elemento(driver, xpath, text: bool=True, retries=3):
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
            if text:
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
            remover_backdrop(driver, False)            # Remove o backdrop antes de inserir texto
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
        remover_backdrop(driver, False)        # Remove o backdrop antes de obter texto
        elemento = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, xpath)))
        texto = elemento.text.strip()
        return texto if texto else None
    except Exception as erro:
        print(f"❌ Erro ao obter texto do elemento {xpath}: {erro}")
        return None


# Função para remover o backdrop
def remover_backdrop(driver, msg=True):
    try:
        driver.execute_script("""
            var backdrops = document.getElementsByClassName('backdrop');
            while(backdrops.length > 0){
                backdrops[0].parentNode.removeChild(backdrops[0]);
            }
        """)
        if msg:
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


# Navega pela primeira página e coleta os dados
def loop_segunda_pagina(driver, source_dir: str, num_parecer: str, codigo: str) -> bool:
    """Acessa página de análises, preenche o formulário e coloca o parecer da proposta como anexo"""

    lista_caminhos = [
        # [0] > Análises
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/ul/li[4]/button',

        # [1] > Adicionar
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-consulta/div['
        '1]/div[1]/button',

        # [2] > Órgão Responsável pela Análise
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-cadastro/form'
        '/div[1]/div/br-select/div/div/div[1]/ng-select/div/div/div[2]/input',

        # [3] > Resultado da Análise
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-cadastro/form'
        '/div[2]/div[1]/br-select/div/div/div[1]/ng-select/div/div/div[2]',

        # [4] > Parecer
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-cadastro/form'
        '/div[4]/div/br-textarea/div/div[1]/div/textarea',

        # [5] > Novo Anexo [0]
        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/transferencia-especial-main'
        '/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/transferencia-cadastro/form'
        '/transferencia-anexo/div/div[1]/button',

        # [6] > Descrição do Arquivo [1]
        '/html/body/modal-container/div[2]/div/div[2]/div/form/div[1]/div/br-input/div/div[1]/input',

        # [7] > Botão Anexo [2]
        '/html/body/modal-container/div[2]/div/div[2]/div/form/div[2]/div/br-file/div/button',

        # [8] > Botão incluir [3]
        '/html/body/modal-container/div[2]/div/div[3]/button[2]'
    ]


    texto_parecer = 'Parecer aprovado com ressalvas conforme anexo.'
    texto_desc_arq = 'Parecer MESP'
    texto_org = '308797 - Ministério do Esporte'
    texto_res_anlz = 'Aprovar Plano de Trabalho'

    print(f"\n{'<' * 20}🤖 Processando parecer: {num_parecer} {'>' * 20}\n".center(50))
    # Aba Análises
    clicar_elemento(driver, lista_caminhos[0])
    remover_backdrop(driver, False)

    # Adiciona nova análise
    clicar_elemento(driver, lista_caminhos[1])
    remover_backdrop(driver, False)

    # Seleciona órgão responsável
    try:
        # Locate and click the dropdown to open it
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, lista_caminhos[2])))
        drop_org = driver.find_element(By.CSS_SELECTOR, "ng-select.brx-input")
        drop_org.click()

        # Find the input field inside the dropdown and type your search text
        search_input = WebDriverWait(drop_org, 10).until(
         EC.visibility_of_element_located((By.CSS_SELECTOR, "ng-select.brx-input input"))
        )
        search_input.send_keys(texto_org)  # Type the exact option text you want

        # Wait for options to appear and click the matching one
        option = WebDriverWait(drop_org, 20).until(
         EC.element_to_be_clickable(
             (By.XPATH, "//div[@class='ng-option']//span[contains(text(), 'Ministério do Esporte')]"))
        )
        option.click()
    except Exception as e:
        print(f'⚠️ Falha ao selecionar órgão responsável na aba de análise.\n Erro{type(e).__name__}\nmsg:'
              f'{str(e)[:100]}')
        discard = WebDriverWait(driver, 3).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "swal2-popup")))
        discard_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'swal2-confirm')]"))
        )
        discard_btn.click()
        if discard:
            save_prop_with_pop(source_dir, codigo)
        return False

    # Resultado da Análise
    try:
        drop_res = driver.find_element(By.XPATH, lista_caminhos[3])
        drop_res.click()
        # Find the input field inside the dropdown and type your search text
        srh_input_res = WebDriverWait(drop_res, 10).until(
            EC.visibility_of_element_located((By.CSS_SELECTOR, "ng-select.brx-input input"))
        )
        srh_input_res.send_keys(texto_res_anlz)  # Type the exact option text you want

        # Wait for options to appear and click the matching one
        opt_res = WebDriverWait(drop_res, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//div[@class='ng-option ng-option-marked']//span[contains(text(),"
                           " 'Aprovar Plano de Trabalho')]"))
        )
        opt_res.click()
    except Exception as e:
        print(f'⚠️ Falha ao colocar resultado da Análise na aba de análise.\n Erro{type(e).__name__}\nmsg:'
              f'{str(e)[:100]}')
        discard = WebDriverWait(driver, 3).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "swal2-popup")))
        discard_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'swal2-confirm')]"))
        )
        discard_btn.click()
        if discard:
            save_prop_with_pop(source_dir, codigo)
        return False

    # Coloca o texto de parecer
    try:
        parecer = driver.find_element(By.XPATH, lista_caminhos[4])
        parecer.clear()
        parecer.send_keys(texto_parecer)

    except Exception as e:
        print(f'⚠️ Falha ao colocar o texto de parecer na aba de análise.\n Erro{type(e).__name__}\nmsg:'
              f'{str(e)[:100]}')
        discard = WebDriverWait(driver, 3).until(
            EC.visibility_of_element_located((By.CLASS_NAME, "swal2-popup")))
        discard_btn = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'swal2-confirm')]"))
        )
        discard_btn.click()
        if discard:
            save_prop_with_pop(source_dir, codigo)
        return False

    # Anexa arquivo parecer
    try:
        if anexar_parecer(
                driver=driver,
                path_list=lista_caminhos[5:],
                source_dir=source_dir,
                texto_desc_arq=texto_desc_arq,
                num_parecer=num_parecer
                ):
            return True

        return False

    except Exception as e:
        print(f'⚠️ Falha fatal {type(e).__name__} ao anexar arquivo na aba de análise.\nErro{str(e)[:100]}\n')
        discard = discar_butn(driver)
        if discard:
            save_prop_with_pop(source_dir, codigo)
        return False


# Clicks discard button
def discar_butn(driver):
    discard_btn = WebDriverWait(driver, 5).until(
        EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'swal2-confirm')]"))
    )
    discard_btn.click()
    print('discarted !')
    return True


def anexar_parecer(driver, path_list: list, source_dir: str, texto_desc_arq: str, num_parecer: str) -> bool:
    try:
        # Clica para adicionar novo anexo
        clicar_elemento(driver,path_list[0])
        remover_backdrop(driver, False)
        time.sleep(1)
        # Coloca o testo do novo arquivo
        desc_arq = driver.find_element(By.XPATH, path_list[1])
        desc_arq.click()
        desc_arq.send_keys(texto_desc_arq)

        # Seleciona qual arquivo enviar
        aqr_path = select_files_by_suffix(source_dir=source_dir,num_parecer=num_parecer)

        # Send the absolute file path to the input
        upload_input = driver.find_element(By.CSS_SELECTOR, "input.upload-input[type='file']")
        print(upload_input.text)
        upload_input.send_keys(aqr_path)

        # Inclui anexo
        clicar_elemento(driver, path_list[3])
        return True

    except TimeoutException or NoSuchElementException or StaleElementReferenceException as erro:
        print(f"❌ Erro ao enviar arquivos: {erro[:80]}")
    except Exception as erro:
        print(f"❌ Erro fatal ao enviar arquivos: {type(erro).__name__}\n{erro[:80]}")
        return False

# saves the proposal that promped a pop-up due to it been already filled
def save_prop_with_pop(source_dir, codigo):
    """
        Find rows where Column A matches search_value exactly,
        then write to Column C in those rows.

        Args:
            source_dir: Path to Excel file directory
            num_parecer: Value to match in Column A (case sensitive)
        """
    # Read the Excel file
    df = pd.read_excel(source_dir)
    # Find exact matches in Column A
    col_a = df.iloc[: , 0]
    match_loc = col_a[col_a == codigo]
    # Write to Column C in matching rows
    if not match_loc.empy:
        row_idx = match_loc.index
        # Write data to Column C (index 2) in these rows
        df.iloc[row_idx, 2] = 'Feito'
        # Save back to Excel
        df.to_excel(source_dir, index=False, dtype=str)
        print(f"Successfully wrote to row: {row_idx+2}")

        return True
    else:
        print("No matches found in Column A")
        return False


def select_files_by_suffix(source_dir, num_parecer):
    """
    Returns the absolute path of the first file in 'directory' whose name after the last underscore matches
    'suffix'.
    If no file matches, returns None.
    """
    print('🔍 Searching for files...')
    for filename in os.listdir(source_dir):
        full_path = os.path.join(source_dir, filename)
        # Skip directories
        if not os.path.isfile(full_path):
            continue
        base = os.path.splitext(filename)[0]
        parts = base.split('_')[-1]
        sub_parts = parts.split('.')[0]
        if sub_parts == num_parecer:
            return full_path  # Return immediately on first match
    return None


# Função principal
def main():
    driver = conectar_navegador_existente()

    planilha_final = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
               r'Social\Teste001\Sofia\Pareceres_SEi\processos_DB - Copia.xlsx')

    source_dir = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                  r'Social\Teste001\Sofia\Pareceres_SEi')

    try:
        df = pd.read_excel(planilha_final)
        df.columns = [col.strip() for col in df.columns]
        print(f"✅ Planilha lida com {len(df)} linhas.")

        if df["Código do Plano de Ação"].duplicated().any():
            print("⚠️ Aviso: Há códigos duplicados na planilha. Removendo duplicatas...")
            df = df.drop_duplicates(subset=["Código do Plano de Ação"], keep="first")
            df.to_excel(planilha_final, index=False)

        # Clica no que menu de navegação
        clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/br-header/"
                                "header/div/div[2]/div/div[1]/button/span")
        # Clica em 'Plano de Ação'
        clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/"
                                "div/br-side-menu/nav/div[3]/a/span[2]")

        remover_backdrop(driver, False)

        # Processa cada linha do DataFrame usando o índice
        for index, row in df.iterrows():
            if not pd.isna(row["Já Feito"]):
                print(f'📝⏭️ Proposta com parecer enviado. Pulando linha {index}...')
                continue
            codigo = str(row["Código do Plano de Ação"])  # Garante que o código seja string
            num_parecer = str(row["Parecer"])#.split('(')[0].strip()  # Garante que o código seja string

            print(f"\n⚙️ Processando código: {codigo} (índice: {index})\n")

            # Clica no ícone para filtrar
            clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/div/div/div/"
                                    "main/transferencia-especial-main/transferencia-plano-acao/"
                                    "transferencia-plano-acao-consulta/br-table/div/div/div/button/i")

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
                                           "div/button[2]", text=False):

                        raise Exception("Falha ao aplicar o filtro")

                    # Aguarda a tabela atualizar
                    WebDriverWait(driver, 7).until(
                        EC.presence_of_element_located((By.XPATH, codigo_tabela_xpath))
                    )
                    time.sleep(1)  # Pequeno delay adicional para garantir a atualização

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

                # Navega até "Plano de Trabalho"
                remover_backdrop(driver, False)
                if not clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/"
                                               "div/div/div/main/transferencia-especial-main/"
                                               "transferencia-plano-acao/transferencia-cadastro/br-tab-set/"
                                               "div/nav/ul/li[3]/button"):
                    raise Exception("Falha ao navegar para 'Plano de Trabalho'")

                # Coleta os dados da segunda pagina
                if loop_segunda_pagina(driver=driver,
                                       source_dir=source_dir,
                                       num_parecer=num_parecer,
                                       codigo=codigo):
                    df.iloc[index, 2] = "Sim"  # Column index 2 is 'C'
                    df.to_excel(planilha_final, index=False)
                # Sobe para o topo da página
                driver.execute_script("window.scrollTo(0, 0);")
                # Clica em "Filtrar" para o próximo código
                clicar_elemento(driver,
                                "/html/body/transferencia-especial-root/br-main-layout/div/div/div/main"
                                "/div[1]/br-breadcrumbs/div/ul/li[2]/a")

            except Exception as erro:
                last_error = truncate_error(f"Main loop element intercepted: {str(erro)}",
                                            150)
                print(f"⚠️ {last_error}")

                try:
                    clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/"
                                            "div/div/div/main/transferencia-especial-main/"
                                            "transferencia-plano-acao/transferencia-plano-acao-consulta/"
                                            "br-table/div/div/div/button", False, 1)
                except:
                    print("⚠️ Falha ao recuperar a tela de consulta. Reiniciando navegação...")
                    # Reseta para pagaina inicial via URL
                    driver.get("https://especiais.transferegov.sistema.gov.br/transferencia-especial"
                               "/programa/consulta")
                    # Clica no que menu de navegação
                    clicar_elemento(driver, "/html/body/transferencia-especial-root/br-main-layout/br-header/"
                                            "header/div/div[2]/div/div[1]/button/span")
                    # Clica em 'Plano de Ação'
                    clicar_elemento(driver,
                                    "/html/body/transferencia-especial-root/br-main-layout/div/div/div/"
                                    "div/br-side-menu/nav/div[3]/a/span[2]")

                    remover_backdrop(driver, False)
                continue
        print(f"✅ Todos os pareceres foram enviados com sucesso")

    except Exception as erro:
        last_error = truncate_error(f"Element intercepted: {str(erro)}")
        print(f"❌ {last_error}")
        print(f'{type(erro).__name__}')


if __name__ == "__main__":
    main()


