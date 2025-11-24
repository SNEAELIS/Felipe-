import re
import time
import sys
import os
import pandas as pd

from datetime import datetime

from webdriver_manager.chrome import ChromeDriverManager

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (TimeoutException, NoSuchElementException,
                                        ElementClickInterceptedException, WebDriverException,
                                        StaleElementReferenceException, ElementNotInteractableException)
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys



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


# Função para obter texto de um elemento
def obter_texto(driver):
    """Access the divInformacao element in the nested frames"""
    try:
        # Switch to default content first (safety measure)
        driver.switch_to.default_content()

        # Step 1: Switch to the main visualization frame
        wait = WebDriverWait(driver, 10)
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "ifrVisualizacao")))

        # Step 2: Now find the element within this frame
        element = wait.until(EC.presence_of_element_located((By.ID, "divInformacao")))

        # Get the text content
        text = element.text.strip()
        print(f"\nElement text: {text}\n")

        return str(text)

    except Exception as e:
        print(f"Error accessing element: {e}")
        return None, None
    finally:
        # Always switch back to default content when done
        driver.switch_to.default_content()


# Função que aplica a espera pelo elemento
def wait_for_element(driver, xpath: str, timeout: int = 5) -> bool:
    """Aguarda até que um elemento esteja presente na página."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath)))
        return True
    except Exception:
        return False


def dicinario_dinamico(data_stream: str, res_list: list, processo_value=None, single=True):
    """
    Create dictionary with dynamic keys for a SINGLE row
    Includes 'Processo' column with provided value

    Args:
        data_stream: A SINGLE string separated by '/' (not a list)
        processo_value: A SINGLE processo value for this row
        single:
        res_list:
    """
    data_frame = {
        'Processo': processo_value,
    }

    if single:
        data_frame['column_1'] = data_stream

    else:
        pattern = r'SNEAELIS/([^(\n]+)'
        matches = re.findall(pattern, data_stream)
        offices = [office for match in matches for office in match.split('/')]

        for j, part in enumerate(offices):
            col_name = f'column_{j + 1}'
            data_frame[col_name] = part

    res_list.append(data_frame)
    return res_list


# Sobe os dados para o arquivo excel
def atualiza_excel(df_list=None, filename=None):
    """
    Update an existing Excel file with new data

    Args:
        df_list: List with data to update (list of rows)
        filename: Path to existing Excel file
    """
    # Check if file exists
    if os.path.exists(filename) and df_list:
        try:
            df_new = pd.DataFrame(df_list)

            print(f"Processing {len(df_new)} rows")
        except Exception as e:
            print(f"❌ Erro ao processar dados: {type(e).__name__}\nError: {str(e)[:200]}...")

        try:
            sheet_names = pd.ExcelFile(filename).sheet_names

            # Read existing file
            with (pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as
                  writer):
                df_new.to_excel(writer, index=False, sheet_name=sheet_names[0])
            print("✅ Todos os dados foram coletados e salvos na planilha!")

        except Exception as e:
            print(f"❌ Erro ao atualizar data_base: {type(e).__name__}\nError: {str(e)[:200]}...")


# Verify is the row has data in it
def has_actual_data(x):
    """Returns True if value is non-empty and non-NA"""
    if pd.isna(x):
        return False
    if isinstance(x, str):
        return x.strip() != ""
    return True  # Numbers, booleans, etc.


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
            elemento = WebDriverWait(driver, 7).until(
                EC.presence_of_element_located((By.XPATH, xpath)))
            elemento.clear()
            elemento.send_keys(texto)
            elemento.send_keys(Keys.ENTER)
            return True
        except Exception as erro:
            print(f"❌ Erro ao inserir texto no campo {xpath} (tentativa {tentativa + 1}): {erro[:50]}")
            time.sleep(1)  # Espera antes de tentar novamente
    print(f"❌ Falha após {retries} tentativas para inserir texto em {xpath}")
    return False


# Função principal
def main():
    planiha_fonte = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                     r'Social\Teste001\Monitoramento processos SEi\Controle SNEAELIS - 2025 - Copia.xlsx')

    planilha_final = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                      r'Social\Teste001\Monitoramento processos SEi\monitoramento_sei.xlsx')


    df = pd.read_excel(planiha_fonte, engine='openpyxl').astype(object)

    pattern = r'^\d+\.\d+/\d+-\d+$'
    filter_mask = df['Processo'].str.match(pattern, na=False)
    filtered_df = df[filter_mask].copy()

    print(f"✅ Planilha lida com {len(filtered_df)} linhas.")

    driver = conectar_navegador_existente()

    try:
        try:
            # Volta para a pagina inicial
            clicar_elemento(driver, '//*[@id="lnkControleProcessos"]/img')
        except:
            print('Já na página inicial')

        result = []

        # Processa cada linha do DataFrame usando o índice
        for idx, row in filtered_df.iterrows():
            print(f'Porcentagem completa: {(idx / len(filtered_df) * 100):.2f}%'.center(50))

            codigo = str(row["Processo"])  # Garante que o código seja string


            if pd.isna(filtered_df.loc[idx, "Processo"]) or codigo == '-' or len(codigo) < 18:
                print(f"Skipping invalid process: {codigo}\n")

                continue

            print(f"\n{'📊' * 3}🔎 PROCESSANDO CÓDIGO: {codigo} (ÍNDICE: {idx}) 🔎{'📊' * 3}\n")

            try:
                # Clica na caixa de texto para pesquisa rápida
                clicar_elemento(driver, '//*[@id="txtPesquisaRapida"]')
                inserir_texto(driver=driver, xpath='//*[@id="txtPesquisaRapida"]',texto=codigo)

                texto_sei = obter_texto(driver=driver)

                if texto_sei == 'Processo não possui andamentos abertos.':
                    dicinario_dinamico(data_stream=texto_sei,
                                       res_list=result,
                                       processo_value=codigo,
                                       )
                    continue


                dicinario_dinamico(data_stream=texto_sei,
                                              res_list=result,
                                              processo_value=codigo,
                                              single=False)



            except Exception as erro:
                last_error = truncate_error(f"Main loop element intercepted: {str(erro)}")
                exc_type, exc_value, exc_tb = sys.exc_info()
                print(f"Error occurred at line: {exc_tb.tb_lineno}")
                print(f"⚠️ {last_error}")
                continue
        else:
            print(result)
            atualiza_excel(df_list=result,
                           filename=planilha_final,
                           )
            driver.quit()
    except Exception as erro:
        last_error = truncate_error(f"Element intercepted: {str(erro)}, {type(erro).__name__}")
        print(f"❌ {last_error}")



def comprehensive_frame_search(driver, target_element_id="divInformacao"):
    """Search more thoroughly through all frames and nested frames"""
    print("=== COMPREHENSIVE FRAME SEARCH ===")

    # Store original window/frame
    original_window = driver.current_window_handle
    driver.switch_to.default_content()

    def search_in_frame(frame_path="main"):
        """Recursively search in current frame and its subframes"""
        print(f"\n--- Searching in: {frame_path} ---")

        # First, check for our target element in current frame
        try:
            elements = driver.find_elements(By.ID, target_element_id)
            if elements:
                print(f"🎉 FOUND ELEMENT in {frame_path}!")
                for element in elements:
                    print(f"Element: {element}")
                    print(f"Text: {element.text}")
                return True
        except Exception as e:
            print(f"Error searching in {frame_path}: {e}")

        # Then look for sub-frames
        frames = driver.find_elements(By.TAG_NAME, "iframe") + driver.find_elements(By.TAG_NAME, "frame")
        print(f"Found {len(frames)} sub-frames in {frame_path}")

        for i, frame in enumerate(frames):
            try:
                frame_name = frame.get_attribute("name") or frame.get_attribute("id") or f"frame_{i}"
                frame_src = frame.get_attribute("src") or "no-src"
                print(f"  Switching to sub-frame {i}: {frame_name}")

                driver.switch_to.frame(frame)

                # Recursively search in this sub-frame
                found = search_in_frame(f"{frame_path} -> {frame_name}")

                if found:
                    return True

                # Switch back to parent frame
                driver.switch_to.parent_frame()

            except Exception as e:
                print(f"  Error with sub-frame {i}: {e}")
                driver.switch_to.parent_frame()

        return False

    # Start search from default content
    found = search_in_frame("main")

    # Always return to default content
    driver.switch_to.default_content()

    if not found:
        print(f"\n❌ Element '{target_element_id}' not found in any frame")

    return found



if __name__ == "__main__":
    main()