import logging.handlers
import subprocess
import fitz  # PyMuPDF
import io
from PIL import Image
from pdf2image import convert_from_path
import img2pdf
from PIL import Image
import numpy as np
import logging
import io
import pdfplumber
import base64
from PyPDF2 import PdfReader, PdfWriter
import requests
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.common import WebDriverException, TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
from selenium.webdriver import ActionChains
from colorama import Fore, Style
from pathlib import Path
import pandas as pd
import time
import os
import sys
import shutil
import json
import traceback
import fitz  # PyMuPDF
import pyautogui
import re
import zipfile

def concatenate_excel_files(input_folder=None, output_file=None):
    path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
            r'Social\Teste001\tabela Código da Natureza de Despesa.xlsx')
    ref_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\Teste001\atT00001 (3) (1).xlsx')

    df = pd.read_excel(path)
    print("Main DF columns:", df.columns.tolist())  # Debug print

    ref_df = pd.read_excel(ref_path, header=None)
    print("Reference DF columns:", ref_df.columns.tolist())  # Debug print

    ref_df = ref_df.applymap(lambda x: str(x).strip().lower() if pd.notna(x) else '')

    unique_values_col_b = ref_df[1].unique()
    print("Unique values in reference:", unique_values_col_b)  # Debug print

    try:
        # first occurrence index
        tipo_mask = [str(x).strip() == 'tipo' for x in unique_values_col_b]
        unique_idx = np.where(tipo_mask)[0]
        if len(unique_idx) == 0:
            raise ValueError("'TIPO' not found in reference file's column B")
        unique_idx = unique_idx[0]
        unique_values_temp = unique_values_col_b[unique_idx+1:]
    except Exception as e:
        print(f"Error finding 'TIPO' marker: {str(e)}")
        print("Values found:", unique_values_col_b)
        return

    unique_values = []
    for val in unique_values_temp:
       if isinstance(val, (str, int, float)):
            if not val or val.startswith('<built-in method'):
                continue
            if '/' in val:
                unique_values.extend([v.strip() for v in val.split('/') if v.strip()])
            else:
                unique_values.append(val.lower())
    print(f"Unique values {unique_values}")


    desc_col = next((col for col in df.columns if 'descrição' in str(col).lower()), None)

    if not desc_col:
        raise ValueError("Column 'Descrição' not found in the main DataFrame")

    for value in unique_values:
        # Print separator and DataFrame
        print('\n' + '=' * 80 + '\n')  # Black line separator (using = for visibility)
        print(f"Results for value: '{value}'\n")

        safe_value = re.escape(value)
        filtered_df = df[df['Descrição'].str.contains(safe_value, case=False, regex=True, na=False)]

        # Print DataFrame with formatting
        pd.set_option('display.max_rows', None)
        with pd.option_context('display.max_columns', None):
            print(filtered_df.to_string(index=False))
        print('\n' + '=' * 80 + '\n')  # Black line separator


def separate_nuclei():
    def aid_func(filename):
        pattern = r'Chamadas\s(.*?)(?:\.pdf|_compressed\.pdf|\.xlsx\.pdf|_compressed_compressed\.pdf)'
        match = re.search(pattern, filename, re.IGNORECASE)
        if match:
            nucleus_name = match.group(1).strip()
            return nucleus_name.replace('_.', '')
        return None

    dir_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\SNEAELIS - Termo Fomento Inst. Léo Moura\Frequencia 926508')

    for root, dirs, files in os.walk(dir_path):
        if 'Frequencias 910782' in dirs:
            dirs.remove('Frequencias 910782')
        for filename in files:
            if not filename.lower().endswith('.pdf'):
                continue
            nucleus = aid_func(filename)

            if nucleus:
                new_dir_path = os.path.join(root, nucleus)
                if not os.path.exists(new_dir_path):
                    print(f"Creating directory: {new_dir_path}")
                    os.makedirs(new_dir_path)
                source_file = os.path.join(root, filename)
                destination_file = os.path.join(new_dir_path, filename)

                print(f"Moving file: {filename} to {new_dir_path}")
                shutil.move(source_file, destination_file)


def pdf_to_xlsx():
    def write_log(log_file, message):
        """Write messages to log file"""
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(message + '\n')

    def extract_tables_with_pdfplumber(pdf_source):
        """Extract tables using pdfplumber with focus on area below 'Nº'"""
        tables = []

        with pdfplumber.open(pdf_source) as pdf:
            for page in pdf.pages:
                # Find the 'Nº' marker position
                y_start = None
                for word in page.extract_words():
                    if word['text'].strip().upper() == 'Nº':
                        y_start = word['top']
                        break

                if y_start is None:
                    continue  # Skip if no Nº found

                # Define crop area (everything below Nº)
                crop_area = (0, y_start - 5, page.width, page.height)  # x0, top, x1, bottom

                # Extract table from this region
                cropped_page = page.crop(crop_area)
                table = cropped_page.extract_table()

                if table and len(table) > 1:  # Ensure we have at least header + one row
                    tables.append(table)

        return tables

    dir_path = (r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência "
                r"Social\SNEAELIS - Termo Fomento Inst. Léo Moura")

    # Create log file
    log_file = os.path.join(dir_path, f"conversion_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")

    for root, dirs, files in os.walk(dir_path):
        if 'Frequencias 910782' in dirs:
            dirs.remove('Frequencias 910782')
        for filename in files:
            if not filename.lower().endswith('.pdf'):
                continue

            pdf_path = os.path.join(root, filename)
            xlsx_filename = filename.replace('.pdf', '.xlsx')
            xlsx_path = os.path.join(root, xlsx_filename)

            log_message = f"\nProcessing {filename}..."
            print(log_message)
            write_log(log_file, log_message)

            try:
                # Try with original PDF first
                try:
                    tables = extract_tables_with_pdfplumber(pdf_path)
                    if tables:
                        log_message = "Extracted tables with pdfplumber"
                        print(log_message)
                        write_log(log_file, log_message)
                    else:
                        raise ValueError("No tables found in initial extraction")
                except Exception as e:
                    log_message = f"Initial extraction failed, trying cleaned version: {str(e)[:200]}"
                    print(log_message)
                    write_log(log_file, log_message)

                    # Clean PDF in memory
                    pdf_reader = PdfReader(pdf_path)
                    pdf_writer = PdfWriter()

                    for page in pdf_reader.pages:
                        if "/Annots" in page:
                            del page["/Annots"]
                        pdf_writer.add_page(page)

                    # Create in-memory PDF bytes
                    pdf_bytes = io.BytesIO()
                    pdf_writer.write(pdf_bytes)
                    pdf_bytes.seek(0)

                    # Try extraction again with cleaned PDF
                    tables = extract_tables_with_pdfplumber(pdf_bytes)
                    if not tables:
                        raise ValueError("All extraction methods failed")

                # Save to Excel
                with pd.ExcelWriter(xlsx_path) as writer:
                    for i, table in enumerate(tables):
                        try:
                            # Convert table data to DataFrame
                            df = pd.DataFrame(table[1:], columns=table[0])
                            df.to_excel(writer, sheet_name=f"Table_{i + 1}", index=False)
                        except Exception as e:
                            log_message = f"Error saving table {i + 1}: {str(e)[:200]}"
                            print(log_message)
                            write_log(log_file, log_message)
                            continue

                log_message = f"Successfully converted {filename} to {xlsx_filename}"
                print(log_message)
                write_log(log_file, log_message)

            except Exception as e:
                log_message = f"Failed to process {filename}: {type(e).__name__} - {str(e)[:200]}"
                print(log_message)
                write_log(log_file, log_message)
                continue


def html_to_pdf(html_file_path, pdf_file_path=None):
    """
    Convert an HTML file to PDF using Chrome browser.

    Args:
        html_file_path (str): Path to the input HTML file
        pdf_file_path (str): Path for the output PDF file (optional)

    Returns:
        bool: True if conversion was successful, False otherwise
    """
    # If no output path is provided, use the same name with .pdf extension
    if pdf_file_path is None:
        pdf_file_path = os.path.splitext(html_file_path)[0] + '.pdf'

    # Set up Chrome options
    chrome_options = Options()
    chrome_options.add_argument('--headless')  # Run in background
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--window-size=1920,1080')  # Set window size for consistent rendering

    try:
        # Initialize the driver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)

        try:
            # Convert file path to URL (handle Windows paths)
            absolute_path = os.path.abspath(html_file_path)
            # Replace backslashes with forward slashes for Windows
            absolute_path = absolute_path.replace('\\', '/')
            file_url = f"file:///{absolute_path}"

            # Navigate to the HTML file
            driver.get(file_url)

            # Wait for page to load (more reliable than implicit wait for PDF generation)
            time.sleep(2)  # Adjust this based on your content complexity

            # Print to PDF with better options
            print_options = {
                'landscape': False,
                'displayHeaderFooter': False,
                'printBackground': True,
                'preferCSSPageSize': True,
                'marginTop': 0.5,
                'marginBottom': 0.5,
                'marginLeft': 0.5,
                'marginRight': 0.5,
                'pageRanges': '',  # Print all pages
            }

            # Execute the print command
            result = driver.execute_cdp_cmd('Page.printToPDF', print_options)

            # Ensure the output directory exists
            os.makedirs(os.path.dirname(pdf_file_path), exist_ok=True)

            # Save the PDF
            with open(pdf_file_path, 'wb') as f:
                f.write(base64.b64decode(result['data']))

            print(f"✓ Successfully converted: {html_file_path} -> {pdf_file_path}")
            return True

        except Exception as e:
            print(f"✗ Error processing {html_file_path}: {str(e)}")
            return False

        finally:
            driver.quit()

    except Exception as e:
        print(f"✗ Failed to initialize browser for {html_file_path}: {str(e)}")
        return False


def convert_all_html_files(root_directory='.', output_directory=None, pattern=('.html', '.htm')):
    """
    Recursively find and convert all HTML files in a directory tree.

    Args:
        root_directory (str): Root directory to search for HTML files
        output_directory (str): Custom output directory (optional)
        pattern (tuple): File extensions to search for
    """
    success_count = 0
    total_count = 0

    print(f"Searching for HTML files in: {os.path.abspath(root_directory)}")

    for root, dirs, files in os.walk(root_directory):
        for file in files:
            if file.lower().endswith(pattern):
                total_count += 1
                html_file_path = os.path.join(root, file)

                # Determine output PDF path
                if output_directory:
                    # Maintain directory structure in output directory
                    relative_path = os.path.relpath(root, root_directory)
                    pdf_dir = os.path.join(output_directory, relative_path)
                    pdf_file_path = os.path.join(pdf_dir, os.path.splitext(file)[0] + '.pdf')
                else:
                    # Save in same directory as HTML file
                    pdf_file_path = None

                # Convert the file
                if html_to_pdf(html_file_path, pdf_file_path):
                    success_count += 1

    print(f"\n📊 Conversion Summary:")
    print(f"   Total HTML files found: {total_count}")
    print(f"   Successfully converted: {success_count}")
    print(f"   Failed: {total_count - success_count}")

    return success_count, total_count


def split_parecer():
    try:
        df = pd.read_excel(r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                           r'Social\Teste001\Sofia\Pareceres_SEi\processos_DB.xlsx', dtype=str)
        for i, r in df.iterrows():
            partial_txt = r.iloc[1]
            txt = partial_txt.split('º')[-1].strip()
            txt = txt.replace(')', '')
            df.at[i,1] = txt
        df.to_excel(r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                           r'Social\Teste001\Sofia\Pareceres_SEi\processos_DB.xlsx',
                    index=False,)
    except Exception as e:
        print(f"✗ Failed to initialize browser for {type(e).__name__}: {str(e)[:100]}")
        return False



def search_google():
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


            return driver
        except Exception as err:
            # Imprime mensagem de erro se a conexão falhar
            print(f"❌ Erro ao conectar ao navegador existente: {err}")

    # Calculate match score (simplified example)
    def calculate_match_score(our_desc, found_title):
        our_desc = str(our_desc).lower()
        found_title = str(found_title).lower()
        our_keywords = set(our_desc.split())
        found_keywords = set(found_title.split())
        intersection = our_keywords.intersection(found_keywords)
        union = our_keywords.union(found_keywords)
        match_score = len(intersection) / len(union) if union else 0
        return round(match_score, 2)

    drive = conectar_navegador_existente()

    df = pd.read_excel(r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                       r'Social\SNEAELIS - Robô PAD\027517_2025 Matheus Sena\Planilha de Custos - Proposta '
                       r'N°27517_2025 (7).xlsx', dtype=str, header=7)
    results_list = []

    for index, row in df.iterrows():
        if row['TIPO'] in ['Recursos Humanos', 'Serviços', 'Eventos']:
            continue

        product_name = row['ITEM']
        our_description = row['ESPECIFICAÇÃO DO ITEM']

        # Create the Google Shopping search URL
        search_query = product_name.replace(' ', '+')
        search_url = f"https://www.google.com/search?q={search_query}&tbm=shop"

        try:
            # Navigate to the search URL using the controlled browser
            drive.get(search_url)

            # Wait for the shopping results to load. We'll wait for a common element.
            # This is a more stable selector than specific class names.
            WebDriverWait(drive, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "[data-dtld]"))
                # Waits for any shopping result card
            )

            # Let the page settle
            time.sleep(2)

            # Now, find the first result. Google's classes change often, so this might need adjustment.
            # Common selectors for the first item's container:
            # Try these if one fails:
            result_selectors = [
                'search', # ID selector
                'div.sh-dgr__grid-result',  # Older selector
                'div.sh-dgr__content',  # Another common one
                'div[data-dtld]',# Generic data attribute
            ]

            first_result = None
            for selector in result_selectors:
                if selector == 'search':
                    first_result = drive.find_element(By.ID, selector)
                    break
                else:
                    results = drive.find_elements(By.CSS_SELECTOR, selector)
                    if results:
                        first_result = results[0]
                        break

            if first_result:
                # Extract text from within the first result
                # Again, use multiple potential selectors for robustness
                try:
                    title_elem = first_result.find_element(
                        By.CSS_SELECTOR, 'h3, [role="heading"], .tAxDx'
                    )
                    found_title = title_elem.text
                except:
                    found_title = "Title not found"

                try:
                    # Look for price element
                    price_elem = first_result.find_element(By.CSS_SELECTOR,
                                                           'span.a8Pemb, .T14wmb, [aria-label*="$"]')
                    found_price = price_elem.text
                except:
                    found_price = "Price not found"

                try:
                    # Look for store/source element
                    source_elem = first_result.find_element(By.CSS_SELECTOR, 'div.aULzUe, .IuHnof')
                    found_source = source_elem.text
                except:
                    found_source = "Source not found"


                results_list.append({
                    'our_product_name': product_name,
                    'found_title': found_title,
                    'found_price': found_price,
                    'found_source': found_source,
                })
                print(f"Success: {product_name} -> {found_price}")

            else:
                print(f"No results found for: {product_name}")
                results_list.append({
                    'our_product_name': product_name,
                    'found_title': "No results found",
                    'found_price': None,
                    'found_source': None,
                    'match_score': None
                })

        except Exception as e:
            print(f"Failed for {product_name}: {e}")
            results_list.append({
                'our_product_name': product_name,
                'found_title': f"Error: {e}",
                'found_price': None,
                'found_source': None,
                'match_score': None
            })

            # Add a delay to be polite and avoid looking like a bot
        time.sleep(3)

        # Save results
        results_df = pd.DataFrame(results_list)
        results_df.to_excel('google_shopping_selenium_results.xlsx', index=False)
        print("Scraping complete. Results saved.")







if __name__ == "__main__":
    func = int(input("Choose a function: "))


    if func == 1:
        concatenate_excel_files()


    # split the nucleai on the dir into subdirs
    elif func == 2:
        separate_nuclei()


    # .pdf to .xlsx
    elif func == 3:
        pdf_to_xlsx()


    # Convert to .pdf from .html
    elif func == 4:
        root_dir = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                    r'Social\Teste001\Sofia\Pareceres_SEi')
        convert_all_html_files(root_directory=root_dir)


    elif func == 5:
        split_parecer()


    elif func == 6:
        set_txt = set()
        text = '''927336_TERMO DE FOMENTO
927336/TERMO DE FOMENTO
927376_TERMO DE FOMENTO
927376/TERMO DE FOMENTO
927768_TERMO DE FOMENTO
927768/TERMO DE FOMENTO
927770_TERMO DE FOMENTO
927770/TERMO DE FOMENTO
927981_TERMO DE FOMENTO
927981/TERMO DE FOMENTO
928381_TERMO DE FOMENTO
928381/TERMO DE FOMENTO
928412_TERMO DE FOMENTO
928412/TERMO DE FOMENTO
934705_TERMO DE FOMENTO
934705/TERMO DE FOMENTO'''
        for txt in text.split('\n'):
            if '_' in txt:
                set_txt.add(txt.split('_')[0])
            else:
                set_txt.add(txt.split('/')[0])
        print(set_txt)


    # confere os pads com time.sleep
    elif func == 7:
        class Robo:
            def webdriver_element_wait(self, xpath: str):
                """
                        Espera até que um elemento web esteja clicável, usando um tempo limite máximo de 3 segundos.

                        Args:
                            xpath: O seletor XPath do elemento.

                        Returns:
                            O elemento web clicável, ou lança uma exceção TimeoutException se o tempo limite for atingido.

                        Raises:
                            TimeoutException: Se o elemento não estiver clicável dentro do tempo limite.
                        """
                # Cria uma instância de WebDriverWait com o driver e o tempo limite e espera o elemento ser clicável
                try:
                    return WebDriverWait(self.driver, 8).until(
                        EC.element_to_be_clickable((By.XPATH, xpath)))
                except Exception as e:
                    raise e

            # Chama a função do webdriver com wait element to be clickable
            def __init__(self):
                """
                Inicializa o objeto Robo, configurando e iniciando o driver do Chrome.
                """
                try:
                    # Configuração do registro
                    # Inicia as opções do Chrome
                    self.chrome_options = webdriver.ChromeOptions()
                    # Endereço de depuração para conexão com o Chrome
                    self.chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
                    # Inicializa o driver do Chrome com as opções e o gerenciador de drivers
                    self.driver = webdriver.Chrome(
                        service=Service(ChromeDriverManager().install()),
                        options=self.chrome_options)
                    self.driver.switch_to.window(self.driver.window_handles[0])

                    print("✅ Conectado ao navegador existente com sucesso.")
                except WebDriverException as e:
                    # Imprime mensagem de erro se a conexão falhar
                    print(f"❌ Erro ao conectar ao navegador existente: {e}")
                    # Define o driver como None em caso de falha na conexão
                    self.driver = None

            # Navega até a página de busca da proposta
            def consulta_proposta(self):
                """
                       Navega pelas abas do sistema até a página de busca de processos.

                       Esta função clica nas abas principal e secundária para acessar a página
                       onde é possível realizar a busca de processos.
                       """
                # Reseta para página inicial
                try:
                    reset = self.webdriver_element_wait('//*[@id="header"]')
                    if reset:
                        img = reset.find_element(By.XPATH, '//*[@id="logo"]/a/img')
                        action = ActionChains(self.driver)
                        action.move_to_element(img).perform()
                        reset.find_element(By.TAG_NAME, 'a').click()
                        # print(Fore.MAGENTA + "\n✅ Processo resetado com sucesso !")
                except NoSuchElementException:
                    print('Já está na página inicial do transferegov discricionárias.')
                except Exception as e:
                    print(Fore.RED + f'🔄❌ Falha ao resetar.\nErro: {type(e).__name__}\n{str(e)[:50]}')

                # [0] Excução; [1] Consultar Proposta
                xpaths = ['//*[@id="menuPrincipal"]/div[1]/div[3]',
                          '//*[@id="contentMenu"]/div[1]/ul/li[2]/a'
                          ]
                try:
                    for idx in range(len(xpaths)):
                        self.webdriver_element_wait(xpaths[idx]).click()
                    # print(f"{Fore.MAGENTA}✅ Sucesso em acessar a página de busca de processo{Style.RESET_ALL}")
                except Exception as e:
                    print(Fore.RED + f'🔴📄 Instrumento indisponível. \nErro: {e}')
                    sys.exit(1)

            def campo_pesquisa(self, numero_processo):
                try:
                    # Seleciona campo de consulta/pesquisa, insere o número de proposta/instrumento e da ENTER
                    campo_pesquisa = self.webdriver_element_wait('//*[@id="consultarNumeroProposta"]')
                    campo_pesquisa.clear()
                    campo_pesquisa.send_keys(numero_processo)
                    campo_pesquisa.send_keys(Keys.ENTER)

                    # Acessa o item proposta/instrumento
                    acessa_item = self.webdriver_element_wait('//*[@id="tbodyrow"]/tr/td[1]/div/a')
                    acessa_item.click()
                except Exception as e:
                    print(
                        f' Falha ao inserir número de processo no campo de pesquisa. Erro: {type(e).__name__}')


            def nav_plano_act_det(self):
                time.sleep(1)
                try:
                    print(f'🚢 Navegando para o plano de ação detalhado:'.center(50, '-'), '\n')
                    # Aba Plano de trabalho
                    self.webdriver_element_wait('//*[@id="div_997366806"]').click()

                    # Seleciona Plano de Aplicação Detalhado
                    self.webdriver_element_wait('//*[@id="menu_link_997366806_836661414"]/div').click()

                except Exception as e:
                    print(
                        f"❌ Ocorreu um erro ao executar ao pesquisar Plano de Ação Detalhado: {type(e).__name__}"
                        f".\n Erro {str(e)[:50]}")


            # Loop para adicionar PAD da proposta
            def loop_de_pesquisa(self, df=None, numero_processo: str=None, tipo_desp: str=None,
                                 cod_natur_desp: str=None, cnpj_xlsx: str=None,
                                 caminho_arquivo_fonte: str=None):

                # Inicia o processo de consulta do instrumento
                try:

                    # Pesquisa pelo processo
                    self.campo_pesquisa(numero_processo=numero_processo)

                    # Executa pesquisa de anexos
                    self.nav_plano_act_det()
                    tab_path = '/html/body/div[3]/div[14]/div[4]/div/div/div/form/div[1]'
                    WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, tab_path)))
                    #time.sleep(15)
                    self.consulta_proposta()
                    return True

                except TimeoutException as t:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(f'TIMEOUT {str(t)[:50]}')
                    print(f"Error occurred at line: {exc_tb.tb_lineno}")
                    self.consulta_proposta()
                except Exception as erro:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(f"Error occurred at line: {exc_tb.tb_lineno}")
                    print(f'❌ Falha ao tentar incluir documentos. Erro:{type(erro).__name__}\n'
                          f'{str(erro)[:100]}...')
                    self.consulta_proposta()

            # Finds which locator to use
            def find_button_with_retry(self):
                # Define all possible locator strategies and values
                locators = [
                    (By.ID, 'form_submit'),
                    (By.NAME, 'detalharEsclarecimentoConvenioDadosDoEsclarecimentoVoltarForm'),
                    (By.XPATH, '//input[@value="Voltar"]'),
                    (By.XPATH, '//td[@class="FormLinhaBotoes"]/input'),
                    (By.CLASS_NAME, 'FormLinhaBotoes'),  # Will need additional find after
                    (By.XPATH, '//input[contains(@onclick, "setaAcao")]'),
                    (By.XPATH, '//*[@id="form_submit"]')
                ]

                for locator in locators:
                    try:
                        # print(f'Botão localizado com o seletor {locator[0]}')
                        return self.driver.find_element(*locator)
                    except Exception as e:
                        print(f"Falha com localizador {locator}: {str(e)[:80]}")
                        continue

                raise NoSuchElementException("Could not find button using any locator strategy")

        def main() -> None:
            l = ['15657/2024']

            try:
                # Instancia um objeto da classe Robo
                robo = Robo()
                # Extrai dados de colunas específicas do Excel
            except Exception as e:
                print(f"\n‼️ Erro fatal ao iniciar o robô: {e}")
                sys.exit("Parando o programa.")

            # inicia consulta e leva até a página de busca do processo
            robo.consulta_proposta()

            for v in l:
                try:
                    robo.loop_de_pesquisa(numero_processo=v)
                except KeyboardInterrupt:
                    print("Script stopped by user (Ctrl+C). Exiting cleanly.")
                    sys.exit(0)  # Exit gracefully
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(f"Error occurred at line: {exc_tb.tb_lineno}")
                    print(
                        f"❌ Falha ao executar script. Erro: {type(e).__name__}\n{str(e)[:100]}")
                    sys.exit(0)  # Exit gracefully

        main()


    elif func == 8:
        root_dir = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                    r'Social\Automações SNEAELIS\Analise_Custos_Exec_Print')
        no_pdf_folder = []  # Use a list to store the folder paths

        for root, dirs, files in os.walk(root_dir):
            # Check if there are any files with a .pdf extension in the current directory
            has_pdf = any(file.lower().endswith('.pdf') for file in files)
            # If no PDF files are found, add the directory name to the list
            if not has_pdf:
                no_pdf_folder.append(os.path.basename(root))

        print("Subdirectories without PDF files:")
        for folder in no_pdf_folder:
            print(folder)


    # Fake long funtion
    elif func == 9:
        class Robo:
            def __init__(self):
                """
                Inicializa o objeto Robo, configurando e iniciando o driver do Chrome.
                """
                try:
                    # Configuração do registro
                    # self.arquivo_registro = ''
                    # Inicia as opções do Chrome
                    self.chrome_options = webdriver.ChromeOptions()
                    # Endereço de depuração para conexão com o Chrome
                    self.chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
                    # Inicializa o driver do Chrome com as opções e o gerenciador de drivers
                    self.driver = webdriver.Chrome(
                        service=Service(ChromeDriverManager().install()),
                        options=self.chrome_options)
                    self.driver.switch_to.window(self.driver.window_handles[0])

                    print("✅ Conectado ao navegador existente com sucesso.")
                except WebDriverException as e:
                    # Imprime mensagem de erro se a conexão falhar
                    print(f"❌ Erro ao conectar ao navegador existente: {e}")
                    # Define o driver como None em caso de falha na conexão
                    self.driver = None

            # Chama a função do webdriver com wait element to be clickable
            def webdriver_element_wait(self, xpath: str, num_element: int = 1):
                """
                        Espera até que um elemento web esteja clicável, usando um tempo limite máximo de 3 segundos.

                        Args:
                            xpath: O seletor XPath do elemento.
                            num_element: Identificador do npúmero de elementos que se espera serem retornados

                        Returns:
                            O elemento web clicável, ou lança uma exceção TimeoutException se o tempo limite for atingido.

                        Raises:
                            TimeoutException: Se o elemento não estiver clicável dentro do tempo limite.
                        """
                # Cria uma instância de WebDriverWait com o driver e o tempo limite e espera o elemento ser clicável
                if num_element == 1:
                    return WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable((By.XPATH, xpath)))
                else:
                    return WebDriverWait(self.driver, 3).until(
                        lambda driver: [elem for elem in driver.find_elements(By.XPATH, xpath)]
                    )

            # Navega até a página de busca do instrumento ou proposta
            def consulta_instrumento(self):
                """
                       Navega pelas abas do sistema até a página de busca de processos.

                       Esta função clica nas abas principal e secundária para acessar a página
                       onde é possível realizar a busca de processos.
                       """
                # Reseta para página inicial
                try:
                    reset = WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable(
                        (By.XPATH, '//*[@id="header"]')))
                    if reset:
                        img = reset.find_element(By.XPATH, '//*[@id="logo"]/a/img')
                        action = ActionChains(self.driver)
                        action.move_to_element(img).perform()
                        reset.find_element(By.TAG_NAME, 'a').click()
                        print(Fore.MAGENTA + "\n✅ Processo resetado com sucesso !")
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}")
                    print(Fore.RED + f'🔄❌ Falha ao resetar.\nErro: {type(e).__name__}\n{str(e)[:100]}')

                # [0] Excução; [1] Consultar Pré-Instrumento/Instrumento
                xpaths = ['//*[@id="menuPrincipal"]/div[1]/div[4]',
                          '//*[@id="contentMenu"]/div[1]/ul/li[6]/a'
                          ]
                try:
                    for idx in range(len(xpaths)):
                        self.webdriver_element_wait(xpaths[idx]).click()
                    print(f"{Fore.GREEN}✅ Sucesso em acessar a página de busca de processo{Style.RESET_ALL}")
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(Fore.RED + f'🔴📄 Instrumento indisponível. \nErro: {e}{Style.RESET_ALL}')
                    sys.exit(1)

            def campo_pesquisa(self, numero_processo):
                try:
                    # Seleciona campo de consulta/pesquisa, insere o número de proposta/instrumento e da ENTER
                    campo_pesquisa = self.webdriver_element_wait('//*[@id="consultarNumeroConvenio"]')
                    campo_pesquisa.clear()
                    campo_pesquisa.send_keys(numero_processo)
                    campo_pesquisa.send_keys(Keys.ENTER)

                    # Acessa o item proposta/instrumento
                    acessa_item = self.webdriver_element_wait('/html/body/div[3]/div[14]/div[3]/div['
                                                              '3]/table/tbody/tr/td/div/a')
                    acessa_item.click()
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(
                        f' Falha ao inserir número de processo no campo de pesquisa. Erro: {type(e).__name__}')

            def busca_convenio(self):
                print('\n🔁🔍 Executando loop de pesquisa de convênio')
                # Seleciona aba primária, após acessar processo/instrumento. Aba Projeto Básico/Termo de referência
                termo_referencia = self.webdriver_element_wait('/html[1]/body[1]/div[3]/div[14]/'
                                                               'div[1]/div[1]/div[1]/a[4]/div[1]/span[1]/span[1]')
                termo_referencia.click()

                # Aba Projeto Básico/Termo de referência
                aba_termo_referencia = self.webdriver_element_wait('/html[1]/body[1]/div[3]/div[14]/div[1]'
                                                                   '/div[1]/div[2]/a[9]/div[1]/span[1]/span[1]')
                aba_termo_referencia.click()

            def acessa_anexos(self):
                try:
                    # Executa pesquisa de anexos
                    print('🔁📎 Executando pesquisa de anexos')

                    # Aba Plano de trabalho
                    self.webdriver_element_wait('//*[@id="div_997366806"]').click()

                    # Seleciona aba anexos
                    self.webdriver_element_wait('//*[@id="menu_link_997366806_1965609892"]/div').click()

                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(f"Ocorreu um erro ao executar ao pesquisar anexos: {type(e).__name__}"
                          f".\n Erro {e[:50]}")

            # Pesquisa o termo de fomento listado na planilha e executa download e transferência caso exista algúm.
            def loop_de_pesquisa(self, numero_processo: str, caminho_pasta: str, pasta_download: str,
                                 err: list = None,
                                 pg: int = 0):
                """
                    Executa as etapas de pesquisa para um número de processo específico.

                    Esta função realiza uma série de interações automatizadas em uma página web
                    para buscar e baixar anexos relacionados a um processo específico.

                    Passos executados:
                    -------------------
                    1. **Pesquisa pelo número do processo**:
                       - Insere o número do processo no campo de busca e pressiona ENTER.
                    2. **Acessa o item correspondente ao processo**.
                    3. **Navega pelas abas**:
                       - Acessa a aba "Plano de Trabalho".
                       - Acessa a aba "Anexos".
                    4. **Baixa os arquivos PDF**:
                       - Verifica se há anexos disponíveis na "Proposta" e inicia o download.
                       - Volta para a aba principal.
                       - Verifica se há anexos na "Execução" e inicia o download.
                    5. **Gerencia os arquivos baixados**:
                       - Aguarda a finalização dos downloads.
                       - Move os arquivos baixados para a pasta correta.
                    6. **Retorna para a página inicial** para processar o próximo número de processo.

                    Args:
                    -----
                    numero_processo : str
                        O número do processo que será pesquisado.

                    caminho_pasta : str
                        Caminho onde os arquivos baixados serão movidos após o download.

                    pasta_download : str
                        Caminho da pasta onde os arquivos são inicialmente baixados.

                    Tratamento de Erros:
                    --------------------
                    - Se algum elemento não for encontrado, uma mensagem de erro será exibida.
                    - Se não houver lista de anexos, o processo continua sem baixar arquivos.
                    - Se houver falha crítica, a execução do programa é encerrada (`sys.exit(1)`).
                """

                # Faz um zip de todos os arquivos em uma pasta
                def zip_all_files_in_folder(folder_path: str, zip_name: str = None, recursive: bool = False,
                                            delete_after_zip: bool = True) -> str:
                    """
                    Compacta todos os arquivos de uma pasta em um arquivo .zip (dentro da própria pasta)
                    e opcionalmente deleta os originais.

                    Parâmetros
                    ----------
                    folder_path : str
                        Caminho da pasta cujos arquivos serão compactados.
                    zip_name : str, opcional
                        Nome do arquivo zip. Se não informado, usará o nome da pasta.
                    recursive : bool, opcional
                        Se True, inclui subpastas. Padrão: False (apenas arquivos na raiz).
                    delete_after_zip : bool, opcional
                        Se True, deleta os arquivos originais após compactação. Padrão: True.

                    Retorno
                    -------
                    Caminho completo do arquivo .zip criado.
                    """
                    folder = Path(folder_path)
                    if not folder.exists() or not folder.is_dir():
                        raise ValueError(f"Pasta '{folder_path}' não existe ou não é válida.")

                    # Define o caminho do ZIP (dentro da pasta alvo)
                    zip_name = zip_name if zip_name else f"{folder.name}.zip"
                    zip_path = folder / zip_name  # Agora dentro da pasta alvo

                    # Lista de arquivos a serem deletados (se delete_after_zip=True)
                    files_to_delete = []

                    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        if recursive:
                            for root, _, files in os.walk(folder):
                                for file in files:
                                    file_path = Path(root) / file
                                    # Ignora o próprio arquivo ZIP durante a criação
                                    if file_path != zip_path:
                                        arcname = file_path.relative_to(folder)
                                        zipf.write(file_path, arcname)
                                        files_to_delete.append(file_path)
                        else:
                            for file in folder.iterdir():
                                if file.is_file() and file != zip_path:  # Evita incluir o ZIP
                                    zipf.write(file, file.name)
                                    files_to_delete.append(file)

                    # Deleta os arquivos originais (se habilitado)
                    if delete_after_zip:
                        for file in files_to_delete:
                            try:
                                if file.is_file():
                                    file.unlink()  # Deleta arquivo
                            except Exception as e:
                                exc_type, exc_value, exc_tb = sys.exc_info()
                                print(
                                    Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                                print(f"⚠️ Erro ao deletar {file}: {str(e)[:100]}")

                    print(f"✅ Arquivo ZIP criado em: {zip_path}")
                    return str(zip_path)

                # Baixa os PDF's da tabela HTML
                def baixa_pdf_exec(pg, end=False):
                    """
                            Baixa os arquivos PDF presentes em uma tabela HTML.

                            Esta função localiza uma tabela HTML com o ID 'listaAnexos', itera sobre suas linhas,
                            extrai o nome do arquivo e a data de upload, e clica no botão de download para cada arquivo.
                            Em seguida, transfere o arquivo baixado para a pasta especificada.

                            Returns:
                                None
                            """
                    try:
                        # Encontra a tabela de anexos
                        tabela = WebDriverWait(self.driver, 10).until(
                            EC.presence_of_element_located((By.ID, 'listaAnexos')))
                        # Diz quantas páginas tem
                        paginas = self.conta_paginas(tabela)

                        print(f'💾📁 Baixando os arquivos do processo {numero_processo}.')

                        for pagina in range(1, paginas + 1):
                            if end:
                                break
                            try:
                                if pg > pagina:
                                    print('🌀📄 pulando página')
                                    continue

                                # Ensure the error log has a list for the current page
                                while len(err) < pagina:
                                    err.append([])

                                # Check if the last page was the tenth page, this site has a max of 10 pages /block
                                if pagina > 1:
                                    if (pagina - 1) % 10 == 0:
                                        self.driver.find_element(By.XPATH,
                                                                 '//*[@id="listaAnexos"]/span[2]/a[10]').click()

                                    elif (pagina - 1) % 10 != 0:
                                        self.driver.find_element(By.LINK_TEXT, f'{pagina}').click()

                                # Encontra a tabela na página atual
                                tabela = self.driver.find_element(By.ID, 'listaAnexos')
                                # Encontra todas as linhas da tabela atual
                                linhas = tabela.find_elements(By.XPATH, './/tbody/tr')

                                for indice, linha in enumerate(linhas):
                                    # Pagina - 1 is to correct the index, synce pagina starts @ 1 and list idx @ 0
                                    if err[pagina - 1]:
                                        if indice <= err[pagina - 1][-1]:
                                            print(f'⏭️ pulando linha: {indice}')
                                            continue
                                    if end:
                                        break
                                    try:
                                        botao_download = linha.find_element(By.CLASS_NAME, 'buttonLink')
                                        if botao_download:
                                            botao_download.click()
                                    except StaleElementReferenceException:
                                        try:
                                            linha_erro = indice - 1
                                            if linha_erro >= 0 and linha_erro not in err[pagina - 1]:
                                                err[pagina - 1].append(linha_erro)
                                            print(
                                                f"⚠️ StaleElementReferenceException occurred at line: {indice}"
                                                f"::::::{err[pagina - 1]}")

                                            self.driver.back()
                                            self.consulta_instrumento()
                                            end = self.loop_de_pesquisa(numero_processo=numero_processo,
                                                                        caminho_pasta=caminho_pasta,
                                                                        pasta_download=pasta_download,
                                                                        err=err,
                                                                        pg=pagina,
                                                                        )
                                        except Exception as error:
                                            exc_type, exc_value, exc_tb = sys.exc_info()
                                            print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}"
                                                             f"{Style.RESET_ALL}")

                                            error_trace = traceback.format_exc()
                                            print(f'❌ Erro ao pular linha com falha. Erro:'
                                                  f' {type(error).__name__}\nTraceback:\n{error_trace}')
                                    except Exception as error:
                                        exc_type, exc_value, exc_tb = sys.exc_info()
                                        print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}"
                                                         f"{Style.RESET_ALL}")

                                        print(
                                            f"❌ Erro ao processar a linha nº{indice} de termo, erro: {type(error).__name__}")
                                        continue

                            except Exception as error:
                                exc_type, exc_value, exc_tb = sys.exc_info()
                                print(
                                    Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                                print(f"❌ Erro ao buscar nova página em anexos execução: {error}. Err:"
                                      f" {type(error).__name__}")

                    except Exception as error:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                        print(f'❌ Erro de download: Termo {error}. Err: {type(error).__name__}')

                self.campo_pesquisa(numero_processo=numero_processo)

                try:
                    self.acessa_anexos()

                    # Seleciona lista de anexos execução e manda baixar os arquivos
                    try:
                        if not err:
                            err = [[]]
                            pg = 0

                        print('\n🔁📎 Executando pesquisa de anexos execução')

                        time.sleep(0.3)
                        botao_lista_execucao = self.webdriver_element_wait('//tbody//tr//input[2]')

                        if botao_lista_execucao.is_displayed() or botao_lista_execucao.is_enabled():
                            try:
                                # Seleciona lista de anexos execução e acessa a mesma
                                botao_lista_execucao.click()
                            except TimeoutException:
                                print(f"Timeout waiting for element: {'//tbody//tr//input[2]'}")
                            except Exception as e:  # Catch other potential exceptions
                                exc_type, exc_value, exc_tb = sys.exc_info()
                                print(
                                    Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                                print(f"🤷‍♂️❌ Erro ao tentar entar na lista de anexos execução: {e}")

                            # Baixa os anexos da tabela.
                            try:
                                baixa_pdf_exec(pg)
                            except Exception as e:
                                exc_type, exc_value, exc_tb = sys.exc_info()
                                print(
                                    Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                                print(f"❌ Tabela não encontrada.\nErro: {e}")
                        else:
                            # Volta para a aba de consulta (começo do loop) caso não tenha lista de execução
                            self.webdriver_element_wait('/html[1]/body[1]/div[3]/div[2]/div[6]/a[2]').click()

                        # espera os downloads terminarem
                        self.espera_completar_download(pasta_download=pasta_download)
                        # Transfere os arquivos baixados para a pasta com nome do processo referente
                        self.transfere_arquivos(caminho_pasta, pasta_download)
                        # Zipa os arquivos da pasta recém criada
                        zip_all_files_in_folder(folder_path=caminho_pasta, )

                        print(
                            f"\n{Fore.GREEN}✅ Loop de pesquisa concluído para o processo:"
                            f" {numero_processo}{Style.RESET_ALL}\n")

                        return True

                    except Exception as er:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                        print(f'❌ Falha ao acessar documentos de execução.'
                              f'\nErro:{type(er).__name__}, {str(er)[:50]}')
                        self.consulta_instrumento()

                except TimeoutException as t:
                    print(f'TIMEOUT {t[:50]}')
                    self.consulta_instrumento()
                except Exception as erro:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(
                        f'❌ Falha ao acessar documentos de execução. Erro:{type(erro).__name__}\n{str(erro)[:50]}')
                    self.consulta_instrumento()

            # Salva a página do browser em PDF e separa o top
            def print_page(self, pasta_download: str, pdf_path: str, crop_height: int = 280):
                """
                   Função para automatizar o processo de impressão da página atual do navegador Chrome como PDF,
                   salvando o arquivo em um caminho especificado e, em seguida, recortando as primeiras páginas
                   do PDF gerado.

                   Descrição Geral:
                   ----------------
                   Esta função realiza duas operações principais:
                   1. Utiliza automação com pyautogui para imprimir a tela atual do Chrome como PDF, salvando no caminho desejado.
                   2. Após o salvamento, realiza o recorte das primeiras páginas do PDF, removendo uma área do topo de cada página,
                      conforme o valor do parâmetro `crop_height`. O arquivo original é sobrescrito pelo PDF recortado.

                   Parâmetros:
                   -----------
                   pasta_download : str
                       Caminho para a pasta onde o PDF será salvo e manipulado.
                   pdf_path : str
                       Caminho completo (incluindo nome do arquivo) onde o PDF será salvo.
                   crop_height : int, opcional
                       Altura (em pontos) a ser recortada do topo das páginas do PDF. Valor padrão: 150.

                   Funcionamento:
                   --------------
                   1. A função interna `save_chrome_screen_as_pdf` automatiza:
                       - Abrir o diálogo de impressão (Ctrl+P) no Chrome.
                       - Confirmar o destino padrão ("Salvar como PDF").
                       - Salvar o PDF no caminho definido.
                   2. A função interna `crop_pdf`:
                       - Localiza o PDF mais recente na pasta de downloads.
                       - Abre o arquivo e recorta até 2 páginas, removendo a altura definida em `crop_height`.
                       - Salva o PDF recortado em um arquivo temporário e substitui o original.
                   3. Mensagens de sucesso ou erro são exibidas durante o processo.

                   Dependências:
                   -------------
                   - pyautogui: Para automação do teclado.
                   - time: Para atrasos sincronizados.
                   - pathlib.Path: Para manipulação de caminhos.
                   - fitz (PyMuPDF): Para manipulação e recorte de PDFs.
                   - os: Para operações sobre arquivos.

                   Exemplo de Uso:
                   ---------------
                   print_page(
                       pasta_download="/caminho/para/downloads",
                       pdf_path="/caminho/para/downloads/arquivo.pdf",
                       crop_height=150
                   )

                   Observações:
                   ------------
                   - É necessário que a janela do Chrome esteja focada antes da execução.
                   - O destino padrão do diálogo de impressão deve estar como "Salvar como PDF".
                   - O recorte é realizado apenas nas duas primeiras páginas do PDF.
                   """
                driver = self.driver

                # Salva a tela do navegar com PDF
                def save_chrome_screen_as_pdf(max_attempts=3, print_delay=2, save_delay=10):
                    """
                    Automatiza o processo de impressão da aba atual do Chrome para PDF.
                    Pressupõe que o Chrome está em foco e que 'Salvar como PDF' é a opção padrão no diálogo de impressão.

                    Parâmetros
                    ----------
                    delay_before : float
                        Tempo de espera (em segundos) antes de enviar o comando Ctrl+P.
                    delay_after_print : float
                        Tempo de espera (em segundos) para o diálogo de impressão aparecer.
                    delay_after_save : float
                        Tempo de espera (em segundos) para a conclusão do salvamento do arquivo.
                    """

                    attempts = 0
                    while attempts < max_attempts:
                        try:
                            WebDriverWait(driver, timeout=10).until(EC.presence_of_element_located((
                                By.XPATH, '//*[@id="form_submit"]')))

                            # Open print dialog
                            pyautogui.hotkey('ctrl', 'p')
                            time.sleep(print_delay)

                            # Press Enter to accept default 'Salvar como PDF'
                            pyautogui.press('enter')
                            time.sleep(1.5)  # Wait for save dialog

                            # Clear any existing path and type new one
                            pyautogui.hotkey('ctrl', 'a')
                            pyautogui.press('backspace')
                            pyautogui.typewrite(pdf_path)
                            time.sleep(1)
                            pyautogui.press('enter')

                            # Wait for file to actually save
                            time.sleep(save_delay)

                            # Return to previous screen, IMPORTANT uses (alt + <-) hotkey
                            pyautogui.hotkey('alt ', 'left')

                            # Verify file exists before proceeding
                            if Path(pdf_path).exists():
                                print(f"✅ PDF saved to: {pdf_path}")
                                return True


                        except Exception as er:

                            print(f'Attempt {attempts + 1} failed: {str(er)[:80]}')
                            attempts += 1
                            time.sleep(2)  # Wait before retry

                        print(f"❌ Failed to save PDF after multiple attempts.")
                        return False

                # Recorta o topo do PDF para retirar dados pessoais e informações inuteis
                def crop_pdf():
                    """
                        Recorta as primeiras páginas de um arquivo PDF localizado na pasta de downloads e substitui o arquivo original pelo novo PDF recortado.

                        Funcionamento:
                        --------------
                        1. Procura na pasta de downloads o arquivo PDF mais recente.
                        2. Abre o PDF e recorta a área superior de cada página (até um limite de 2 páginas), utilizando o valor de `crop_height` para definir o quanto será removido da borda superior.
                        3. Salva o resultado em um arquivo temporário com prefixo 'cropped_'.
                        4. Substitui o arquivo original pelo arquivo recortado usando `os.replace`, garantindo que o original seja sobrescrito de forma segura.
                        5. Exibe mensagens de sucesso ou erro durante o processo.

                        Parâmetros:
                        -----------
                        Não recebe parâmetros diretamente. Espera que as variáveis globais `pasta_download` (caminho da pasta de download) e `crop_height` (altura a ser recortada do topo das páginas) estejam definidas.

                        Exceções Tratadas:
                        ------------------
                        - FileNotFoundError: Caso não encontre arquivos PDF na pasta de download.
                        - Outros erros de leitura, manipulação ou gravação do PDF são capturados e impressos na tela.

                        Dependências:
                        -------------
                        - pathlib.Path para manipulação de caminhos.
                        - fitz (PyMuPDF) para manipulação de PDFs.
                        - os para operações de sistema de arquivos.

                        Exemplo de Uso:
                        ---------------
                        pasta_download = "/caminho/para/pasta/downloads"
                        crop_height = 150
                        crop_pdf()
                        """
                    try:
                        download_path = Path(pasta_download)
                        pdf_files = list(download_path.glob('*.pdf'))
                        if not pdf_files:
                            raise FileNotFoundError("No PDF files found in downloads")
                        newest_pdf = max(pdf_files, key=os.path.getctime)
                    except Exception as e:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                        print(
                            f"❌ [File Search Error] Could not find or select PDF file: {type(e).__name__}:\n"
                            f"{str(e)[:100]}")
                        return

                    try:
                        page_num = 0
                        doc = fitz.open(newest_pdf)
                        for page in doc:
                            if page_num > 0:
                                break
                            rect = page.rect
                            new_rect = fitz.Rect(
                                rect.x0,
                                rect.y0 + crop_height,
                                rect.x1,
                                rect.y1
                            )
                            page.set_cropbox(new_rect)
                            page_num += 1
                        # Save to a temporary new file
                        temp_pdf = newest_pdf.with_name("cropped_" + newest_pdf.name)
                        doc.save(temp_pdf)
                        doc.close()
                        # Replace the original file
                        os.replace(temp_pdf, newest_pdf)
                        print(f"✅ PDF '{newest_pdf.name}' foi cortado com sucesso (arquivo substituído).")
                    except Exception as e:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                        print(
                            f"❌ [PDF Processing Error] Error during PDF cropping: {type(e)[:100].__name__}: "
                            f"{str(e)[:100]}")
                        return

                # Main execution
                if save_chrome_screen_as_pdf():
                    crop_pdf()

            # Limpa o nome do arquivo
            @staticmethod
            def clean_filename(filename):
                # Remove or replace invalid characters for Windows filenames
                return re.sub(r'[\\/*?:"<>|]', "_", filename)

            # Salva a tela de esclarecimento detalhado.
            def loop_esclarecimento(self, pasta_download: str, caminho_pasta: str, numero_processo: str):
                def baixa_respostas():
                    print(f'💾📁 Baixando os esclarecimentos do processo: {numero_processo}.')
                    try:
                        # Encontra a tabela na página atual
                        tabela_resp = self.driver.find_element(
                            By.XPATH,
                            '/html/body/div[3]/div[14]/div[3]/div/div/form/table/tbody/tr[18]/td/div')
                        # Encontra todas as linhas da tabela atual
                        linhas_resp = tabela_resp.find_elements(By.XPATH, './/tbody/tr')

                        # Itera as linhas da tabela
                        for idx, linha_r in enumerate(linhas_resp):
                            try:
                                botao_download = linha_r.find_element(By.CLASS_NAME, 'buttonLink')
                                if botao_download:
                                    botao_download.click()
                            except NoSuchElementException:
                                continue
                            except Exception as err:
                                exc_type, exc_value, exc_tb = sys.exc_info()
                                print(
                                    Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                                print(
                                    f"❌ Erro ao processar a linha nº{idx} de termo,"
                                    f" erro: {type(err).__name__}")
                                continue
                    except Exception as err:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                        print(f'❌ Erro de download: Termo {str(err)[:100]}. Err: {type(err).__name__}')

                # Xpath [0] Acomp. e Fiscalização // [1] Esclarecimento
                loop_list = ['//*[@id="menuInterno"]/div/div[7]',
                             '//*[@id="contentMenuInterno"]/div/div[1]/ul/li[3]/a']
                for j in loop_list:
                    try:
                        self.webdriver_element_wait(j).click()
                    except TimeoutException as e:
                        print(f'❌ Falha ao acessar documentos de esclarecimento. Erro:{type(e).__name__}')

                print(f'🖨️🖼️ Imprimindo tela do processo {numero_processo}.')

                try:
                    # Encontra a tabela de lista de esclarecimento
                    tabela = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.ID, 'esclarecimentos')))

                    # Encontra todas as linhas da tabela
                    linhas = tabela.find_elements(By.XPATH, './/tbody/tr')

                    # Diz quantas páginas tem
                    paginas = self.conta_paginas(tabela)

                    try:
                        for pagina in range(1, paginas + 1):
                            if pagina > 1:
                                element = WebDriverWait(self.driver, 10).until(
                                    EC.element_to_be_clickable((By.LINK_TEXT, f'{pagina}')))
                                element.click()
                                print(f'\nAcessando página {pagina}\n ')

                                # Encontra a tabela na página atual
                                tabela = WebDriverWait(self.driver, 10).until(
                                    EC.presence_of_element_located((By.ID, 'esclarecimentos')))
                                # Encontra todas as linhas da tabela atual
                                linhas = tabela.find_elements(By.TAG_NAME, 'tr')

                            for indice in range(1, len(linhas) + 1):
                                try:
                                    # Refresh table and rows reference periodically
                                    if indice > 0:
                                        tabela = WebDriverWait(self.driver, 10).until(
                                            EC.presence_of_element_located((By.ID, 'esclarecimentos')))
                                        linhas = tabela.find_elements(By.TAG_NAME, 'tr')
                                except Exception as error:
                                    exc_type, exc_value, exc_tb = sys.exc_info()
                                    print(
                                        Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                                    print(
                                        f"❌ Erro ao processar a linha nº{indice},"
                                        f" erro: {type(error).__name__}\n{str(error)[:80]}")
                                    break

                                # Get fresh reference to the current row
                                linha = linhas[indice]

                                # Get PDF name components
                                pdf_name_1 = linha.find_element(By.CLASS_NAME, 'sequencial').text
                                pdf_name_2 = linha.find_element(By.CLASS_NAME, 'dataSolicitacao').text
                                raw_name = f"{pdf_name_2}_{pdf_name_1}.pdf"

                                # Process filename and path
                                safe_name = self.clean_filename(raw_name)
                                full_pdf_path = os.path.join(pasta_download, safe_name)
                                unique_pdf_path = self.make_unique_path(full_pdf_path)

                                try:
                                    # Handle download button with fresh reference
                                    botao_detalhar = linha.find_element(By.CLASS_NAME, 'buttonLink')
                                    botao_detalhar.click()

                                    # Baixa arquivos da lista de respostas
                                    baixa_respostas()

                                    # Print page and handle subsequent actions
                                    self.print_page(pasta_download=pasta_download, pdf_path=unique_pdf_path, )

                                    # Use explicit wait for the final button
                                    self.driver.execute_script(
                                        "window.scrollTo(0, document.body.scrollHeight);")
                                    pyautogui.hotkey('alt', 'left')

                                except Exception as error:
                                    exc_type, exc_value, exc_tb = sys.exc_info()
                                    print(
                                        Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                                    print(
                                        f'❌ Erro ao imprimir a página. Erro: {type(error).__name__}\n'
                                        f'Traceback:\n{str(error)[:50]}')
                    except Exception as error:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                        print(f"❌ Erro ao buscar nova página de esclarecimento: {error}.\nErr:"
                              f" {type(error).__name__}")
                except Exception as error:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(f"❌ Erro ao econtrar elementos na página de esclarecimentos: {error}. Err:"
                          f" {type(error).__name__}")

                self.espera_completar_download(pasta_download=pasta_download)
                self.transfere_arquivos(caminho_pasta, pasta_download)
                self.consulta_instrumento()

                print(
                    f"\n{Fore.GREEN}✅ Loop de esclarecimento concluído para o processo:"
                    f" {numero_processo}{Style.RESET_ALL}\n")

            # Make sure the name of the document is unique
            @staticmethod
            def make_unique_path(path):
                """
                Receives a file path and returns a unique path by adding a counter or timestamp if necessary.
                """
                base, ext = os.path.splitext(path)
                counter = 1
                unique_path = path
                while os.path.exists(unique_path):
                    # Option 1: With counter
                    unique_path = f"{base}_{counter}{ext}"
                    counter += 1
                    # Option 2: With timestamp (uncomment if you prefer this)
                    # unique_path = f"{base}_{int(time.time())}{ext}"
                return unique_path

            # Finds which locator to use
            def find_button_with_retry(self):
                # Define all possible locator strategies and values
                locators = [
                    (By.ID, 'form_submit'),
                    (By.NAME, 'detalharEsclarecimentoConvenioDadosDoEsclarecimentoVoltarForm'),
                    (By.XPATH, '//input[@value="Voltar"]'),
                    (By.XPATH, '//td[@class="FormLinhaBotoes"]/input'),
                    (By.CLASS_NAME, 'FormLinhaBotoes'),  # Will need additional find after
                    (By.XPATH, '//input[contains(@onclick, "setaAcao")]'),
                    (By.XPATH, '//*[@id="form_submit"]')
                ]

                for locator in locators:
                    try:
                        print(f'Botão localizado com o seletor {locator[0]}')
                        return self.driver.find_element(*locator)
                    except Exception as e:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                        print(f"Falha com localizador {locator}: {str(e)[:100]}")
                        continue

                raise NoSuchElementException("Could not find button using any locator strategy")

            # Pega o ultimo arquivo baixado da pasta Downloads e move para a pasta destino
            def espera_completar_download(self, pasta_download: str, tempo_limite: int = 30,
                                          extensoes_temporarias: list = None):
                """
                Aguarda a conclusão de um download verificando a ausência de arquivos temporários.

                Parâmetros:
                -----------
                pasta_download : str
                    Caminho da pasta onde os arquivos estão sendo baixados.

                tempo_limite : int, opcional (padrão: 30)
                    Tempo máximo (em segundos) que a função aguardará antes de levantar uma exceção.

                extensoes_temporarias : list, opcional (padrão: ['.crdownload', '.part', '.tmp'])
                    Lista de extensões de arquivos temporários que indicam download em progresso.

                Retorna:
                --------
                bool
                    Retorna `True` assim que todos os downloads forem concluídos (quando não houver mais arquivos
                    com as extensões temporárias na pasta).

                Levanta:
                --------
                Exception
                    Se o tempo limite for atingido e ainda houver arquivos temporários, uma exceção é lançada.

                Descrição:
                ----------
                - A função inicia um temporizador que expira após `tempo_limite` segundos.
                - Em um loop contínuo, verifica se existem arquivos com as extensões temporárias na pasta de downloads.
                - Se não houver arquivos temporários, retorna `True`, indicando que o download foi concluído.
                - Caso contrário, aguarda 1 segundo antes de verificar novamente.
                - Se o tempo limite for atingido e ainda existirem arquivos temporários, uma exceção é lançada.
                """
                # Define extensões padrão se não fornecidas
                if extensoes_temporarias is None:
                    extensoes_temporarias = ['.crdownload', '.part', '.tmp', '.download']

                # Inicia comparador de tempo
                tempo_final = time.time() + tempo_limite

                while time.time() < tempo_final:
                    # Verifica se há arquivos com qualquer uma das extensões temporárias
                    arquivos_temporarios = [
                        f for f in os.listdir(pasta_download)
                        if any(f.endswith(ext) for ext in extensoes_temporarias)
                    ]

                    time.sleep(1)

                    if not arquivos_temporarios:
                        return True

                raise Exception(
                    f'Não completou o download no tempo limite.'
                    f' Arquivos temporários encontrados: {arquivos_temporarios}')

            # Conta quantas páginas tem para iterar sobre
            def conta_paginas(self, tabela):
                # Diz quantas páginas tem
                try:
                    paginas = tabela.find_element(By.TAG_NAME, 'span').text
                    paginas = paginas.split('(')[0]
                    paginas.strip()
                    paginas = int(paginas[-3:])
                    return paginas
                except NoSuchElementException:
                    paginas = 1
                    return paginas

            def converter_data(self, data: str) -> datetime | None:
                """
                Função auxiliar para converter uma string de data em um objeto datetime,
                tratando diferentes formatos.

                Args:
                    data: String de data a ser convertida.

                Returns:
                    Um objeto datetime correspondente à data fornecida ou None se o formato for inválido.
                """
                formatos_tentativa = [
                    '%Y-%m-%d %H:%M:%S',  # Formato com hora
                    '%Y-%m-%d',  # Formato sem hora
                    '%Y-%m-%dT%H:%M:%S.%fZ',  # Formato ISO 8601 com microssegundos e 'Z'
                    '%d/%m/%Y'  # Formato de data BR
                ]

                for formato in formatos_tentativa:
                    try:
                        return datetime.strptime(data, formato)
                    except ValueError:
                        pass  # Tenta o próximo formato

                print(f"⚠️📆 Formato de data inválido: {data}")
                return None

            def data_hoje(self):
                # pega a data do dia que está executando o programa
                agora = datetime.now()
                data_hoje = datetime(agora.year, agora.month, agora.day)
                return data_hoje

            # Permuta os pontos e traços do nome do arquivo por
            def permuta_nome_arq(self, pasta_download, extension_change=False):
                """
                    Renomeia arquivos em uma pasta removendo espaços do nome ou encurtando nomes longos.

                    Esta função percorre todos os arquivos na pasta especificada e realiza as seguintes alterações:
                    - Remove espaços em branco do nome do arquivo.
                    - Se o nome do arquivo for muito longo (mais de 80 caracteres), ele é renomeado para `Documento (X).ext`, onde `X` é um contador.
                    - Garante que apenas arquivos (e não diretórios) sejam processados.
                    - Ignora arquivos ocultos (aqueles que começam com `.`).

                    Args:
                    -----
                    pasta_download : str
                        Caminho da pasta onde os arquivos serão renomeados.

                    extension_change : bool, optional
                        Parâmetro reservado para futuras modificações na extensão dos arquivos (padrão: `False`).

                    Retorna:
                    --------
                    None

                    Exceções Tratadas:
                    ------------------
                    - Se a pasta não for encontrada, um erro será exibido e a função encerrada.
                    - Se ocorrer um erro ao renomear um arquivo, a mensagem de erro será exibida e a função encerrada.
                """
                if not os.path.isdir(pasta_download):
                    print(f"Error: Folder '{pasta_download}' not found.")
                    return None
                # Contador para renomear os arquivos
                cont = 1

                for nome_arq in os.listdir(pasta_download):
                    # Pula arquivos ocultos
                    if nome_arq.startswith('.'):
                        continue
                    old_filepath = os.path.join(pasta_download, nome_arq)

                    # Lê apenas arquivos, não lê subdiretórios
                    if os.path.isfile(old_filepath):
                        base, ext = nome_arq.rsplit('.', 1)

                        # Verirfica o tamanho do nome do arquivo para que não dê erro de transferência
                        if len(base) > 80:
                            novo_nome = f'Documento ({cont})'
                            novo_nome = novo_nome + '.' + ext
                            new_filepath = os.path.join(pasta_download, novo_nome)
                            try:
                                cont += 1
                                os.rename(old_filepath, new_filepath)
                            except OSError as e:
                                print(f"Error renaming '{nome_arq}': {e}")
                                return None  # Or handle the error differently
                        else:
                            novo_nome = base.replace(' ', '').strip()
                            novo_nome = novo_nome + '.' + ext
                            new_filepath = os.path.join(pasta_download, novo_nome)
                            try:
                                os.rename(old_filepath, new_filepath)
                            except OSError as e:
                                print(f"Error renaming '{nome_arq}': {e}")
                                return None  # Or handle the error differently

            # Coloca o arquivo na pasta do OneDrive
            def transfere_arquivos(self, caminho_pasta: str, pasta_download: str, tempo_limite=10):
                """
                    Move arquivos recém-baixados da pasta de downloads para um diretório de destino.

                    Parâmetros:
                    -----------
                    pasta_download : str
                        Caminho da pasta onde os arquivos foram baixados.

                    caminho_pasta : str
                        Caminho da pasta de destino para onde os arquivos serão movidos.

                    Comportamento:
                    --------------
                    - Obtém a data de hoje usando `self.data_hoje()`.
                    - Percorre todos os arquivos na `pasta_download`.
                    - Ignora qualquer diretório dentro da pasta.
                    - Verifica a data de modificação do arquivo (`getmtime`).
                    - Se o arquivo foi modificado **hoje**, move-o para `caminho_pasta`.
                    - Exibe uma mensagem no console para cada arquivo movido.
                    """
                # pega a data do dia que está executando o programa
                data_hoje = self.data_hoje()
                tempo_final = time.time() + tempo_limite
                self.permuta_nome_arq(pasta_download)
                moved_files = 0
                while time.time() < tempo_final:
                    if not os.listdir(pasta_download):
                        break
                    try:
                        for arq in os.listdir(pasta_download):
                            caminho_arq = os.path.join(pasta_download, arq)
                            # Data de modificação do arquivo
                            data_mod = datetime.fromtimestamp(os.path.getmtime(caminho_arq))

                            # Compara a data de hoje com a data de modificação do arquivo
                            if data_mod >= data_hoje:
                                # Move o arquivo para a pasta destino
                                shutil.move(caminho_arq, os.path.join(caminho_pasta, arq))
                                moved_files += 1
                    except Exception as e:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                        print(f'❌ Falha ao mover arquivo {e}')

                print(f'📂 Total de arquivos movidos: {moved_files}')

            # Cria uma pasta com o nome especificado no one-drive e retorna o caminho.
            def criar_pasta(self, nome_pasta: str, caminho_pasta_onedrive: str, tipo_instrumento: str) -> str:
                """Cria uma pasta com o nome especificado no caminho do OneDrive.

                    Substitui caracteres '/' por '_' no nome da pasta para evitar erros.

                    Args:
                        nome_pasta: O número da proposta (nome da pasta a ser criada).
                        caminho_pasta_onedrive: O caminho base para a pasta do OneDrive.
                        tipo_instrumento: Identificador de tipo de consulta executada.

                    Returns:
                        O caminho completo da pasta criada.

                    Raises:
                        Exception: Se ocorrer um erro durante a criação da pasta.
                """
                # Combina o caminho base do OneDrive com o nome da pasta, substituindo '/' por '_'
                nome_pasta = nome_pasta + '_' + tipo_instrumento
                caminho_pasta = os.path.join(caminho_pasta_onedrive, nome_pasta.replace('/', '_'))

                try:
                    # Cria o diretório, incluindo o diretório pai, se necessário.
                    os.makedirs(caminho_pasta, exist_ok=True)
                    print(f"✅ Pasta '{nome_pasta}' criada em: {caminho_pasta}\n")
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(f"❌ Erro ao criar a pasta '{nome_pasta}': {e}")
                # Retorna o caminho completo da pasta, mesmo que a criação tenha falhado (para tratamento posterior)
                return caminho_pasta

            def extrair_dados_excel(self, caminho_arquivo_fonte, busca_id: str, tipo_instrumento_id: str,
                                    situacional_id: str) -> tuple:
                """
                   Lê os contatos de uma planilha Excel e executa ações baseadas nos dados extraídos.

                   Args:
                       caminho_arquivo_fonte (str): Caminho do arquivo Excel que será lido.
                       busca_id (str): Nome da coluna que contém os números de processo.
                       tipo_instrumento_id (str): Nome da coluna que contém o tipo de processo.
                   """

                pd.set_option('future.no_silent_downcasting', True)
                dados_processo = pd.read_excel(caminho_arquivo_fonte, dtype=str)
                dados_processo = dados_processo.fillna('')
                dados_processo = dados_processo.replace({u'\xa0': ''})
                dados_processo = dados_processo.infer_objects(copy=False)

                # Cria um lista para cada coluna do arquivo xlsx
                numero_processo = list()
                tipo_instrumento = list()
                situacional = list()

                try:
                    # Itera a planilha e armazena os dados em listas
                    for indice, linha in dados_processo.iterrows():  # Assume que a primeira linha e um cabeçalho
                        numero_processo.append(linha[busca_id])  # Busca o número do processo
                        tipo_instrumento.append(linha[tipo_instrumento_id])  # Busca o tipo de instrumento
                        situacional.append(linha[situacional_id])  # Busca a situação do processo
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(f"❌ Erro de leitura encontrado, erro: {e}")

                return numero_processo, tipo_instrumento, situacional

            # Verifica se a pasta está vazia
            def pasta_vazia(self, pasta_pai: str) -> list:
                """
                Identifica todas as pastas vazias dentro de um diretório pai.

                Esta função percorre todas as pastas dentro de um diretório especificado (`pasta_pai`)
                e verifica se elas estão vazias. Caso encontre pastas vazias, elas são adicionadas
                a uma lista, que é retornada ao final da execução.

                Parâmetros:
                    pasta_pai (str): O caminho do diretório pai onde a busca será realizada.

                Retorna:
                    list: Uma lista contendo os caminhos completos das pastas vazias encontradas.

                Observações:
                    - Certifique-se de que o caminho fornecido em `pasta_pai` é válido e acessível.
                    - Apenas pastas diretamente contidas em `pasta_pai` serão verificadas (não verifica subpastas recursivamente).
                """
                pastas_vazias = []

                for pasta in os.listdir(pasta_pai):
                    caminho_pasta = os.path.join(pasta_pai, pasta)
                    if os.path.isdir(caminho_pasta) and not os.listdir(caminho_pasta):
                        pastas_vazias.append(caminho_pasta)

                return pastas_vazias

            # Limpa os dados que vem da planilha
            def limpa_dados(self, lista: list):
                """
                Limpa os elementos de uma lista, removendo quebras de linha, espaços extras e
                substituindo valores NaN por strings vazias.

                :param lista: Lista contendo os elementos a serem processados.
                :return: Uma nova lista com os elementos limpos.
                """
                lista_limpa = []
                for i in lista:
                    if pd.isna(i):
                        lista_limpa.append('')
                        continue
                    # Remove quebra de linha "\n"
                    i_limpo = i.replace('\n', '').replace('\r', '').strip()
                    # Anexa à lista limpa
                    lista_limpa.append(i_limpo)
                return lista_limpa

            # Salva o progresso em um arquivo json
            def salva_progresso(self, arquivo_log: str, processo_visitado: str, arquivos_baixados: list,
                                indice: int):
                """
                Salva o progresso atual em um arquivo JSON.

                :param arquivo_log: Endereço do arquivo JSON
                :param processo_visitado: Lista de processos concluídos.
                :param arquivos_baixados: Lista de arquivos baixados.
                :param indice: Diz qual linha o programa iterou por último.
                """
                # Carrega os dados antigos
                dados_log = self.carrega_progresso(arquivo_log=arquivo_log)

                # Carrega os dados novos
                novo_item = {
                    processo_visitado: arquivos_baixados,
                    "indice": indice
                }

                dados_log.update(novo_item)
                with open(arquivo_log, 'w', encoding='utf-8') as arq:
                    json.dump(dados_log, arq, indent=4)
                print(f"💾 Progresso salvo no processo:{processo_visitado} e ìndice: {dados_log["indice"]}")

            # Carrega os dados do arquivo JSON que sereve como Cartão de Memória
            def carrega_progresso(self, arquivo_log: str):
                """
                    Carrega o progresso do arquivo JSON.

                    :param arquivo_log: Endereço do arquivo JSON

                    :return: Um dicionário contendo os dados de progresso.
                             Se o arquivo não existir, retorna valores padrão.
                """
                with open(arquivo_log, 'r') as arq:
                    return json.load(arq)

            # Reseta o arquivo JSON
            def reset(self, arquivo_log: str):
                """
                Reseta o progresso atual em um arquivo JSON.

                :param arquivo_log: Endereço do arquivo JSON

                """
                dados_vazios = {
                    "processo_visitado": [],
                    "indice": 0
                }
                # Salva os dados vazios no arquivo JSON
                with open(arquivo_log, 'w', encoding='utf-8') as arq:
                    json.dump(dados_vazios, arq, indent=4)

            # Verifica as condições para mandar um e-mail para o técnico
            def condicao_email(self, numero_processo: str, caminho_pasta: str):
                """
                Verifica se há arquivos modificados na data de hoje dentro de um diretório específico.

                :param numero_processo: Número do processo relacionado aos arquivos.
                :param caminho_pasta: Caminho da pasta onde os arquivos estão armazenados.
                :return: Uma tupla contendo (numero_processo, caminho_pasta, lista de arquivos modificados hoje).
                         Caso não haja arquivos modificados hoje, retorna uma lista vazia.
                """

                # lista que guarda os arquivos novos, na função ENVIAR_EMAIL_TECNICO recebe o nome de lista_documentos
                docs_atuais = []
                try:
                    # Data de hoje
                    hoje = self.data_hoje()
                    # Itera os arquivos da pasta para buscar a data de modificação individual
                    for arq_nome in os.listdir(caminho_pasta):
                        arq_caminho = os.path.join(caminho_pasta, arq_nome)
                        # Pula diretórios
                        if os.path.isfile(arq_caminho):
                            data_mod = datetime.fromtimestamp(os.path.getmtime(arq_caminho))
                            # Compara as datas de modificação dos arquivos
                            if data_mod >= hoje:
                                docs_atuais.append(arq_nome)
                    if docs_atuais:
                        print(f"📂✨ Documentos novos encontrados para o processo {numero_processo}")
                        return numero_processo, caminho_pasta, docs_atuais

                    else:

                        print(f"⚠️Nenhum documento novo encontrado para o processo {numero_processo}.\n")
                        return []
                except Exception:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(f"Error occurred at line: {exc_tb.tb_lineno}")
                    print(f"⚠️Nenhum documento novo encontrado para o processo {numero_processo}.\n")
                    return []


        def main() -> None:
            def eta():
                idx = indice + 1
                elapsed_time = time.time() - start_time

                # Média por iteração
                avg_time_per_iter = elapsed_time / idx

                # Estimativa de tempo restante
                remaining_iters = max_linha - idx
                eta_seconds = remaining_iters * avg_time_per_iter

                # Formata ETA como mm:ss
                eta_minutes = int(eta_seconds // 60)
                eta_secs = int(eta_seconds % 60)

                print(
                    f"\n{indice} {'>' * 10} Porcentagem concluída:"
                    f" {(indice / max_linha) * 100:.2f}% | ETA: {eta_minutes:02d}:{eta_secs:02d}\n")

            # Caminho da pasta download que é o diretório padrão para download. Use o caminho da pasta 'Download' do
            # seu computador
            pasta_download = r'C:\Users\felipe.rsouza\Downloads'

            # Caminho do arquivo .xlsx que contem os dados necessários para rodar o robô
            caminho_arquivo_fonte = (
                r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\Documentos\awfawf\source\RTMA Passivo 2024 - '
                r'PROJETOS E PROGRAMAS (3) Atualizado em Julho de.2025.xlsx ')
            # Rota da pasta onde os arquivos baixados serão alocados, cada processo terá uma subpasta dentro desta
            caminho_pasta_onedrive = (
                r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\Documentos\awfawf')
            # Caminho do arquivo JSON que serve como catão de memória
            arquivo_log = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                           r'Social\Documentos\awfawf\source\arquivo_log.json')

            try:
                # Instancia um objeto da classe Robo
                robo = Robo()
                # Extrai dados de colunas específicas do Excel
            except Exception as e:
                exc_type, exc_value, exc_tb = sys.exc_info()
                print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                print(f"\n‼️ Erro fatal ao iniciar o robô: {e}")
                sys.exit("Parando o programa.")

            numero_processo, tipo_instrumento, situacional = robo.extrair_dados_excel(
                caminho_arquivo_fonte=caminho_arquivo_fonte,
                busca_id='Instrumento nº',
                tipo_instrumento_id='Regime Jurídico do Instrumento (Modalidade)',
                situacional_id='SITUACIONAL'
            )

            max_linha = len(numero_processo)

            # Pós-processamento dos dados para não haver erros na execução do programa
            numero_processo = robo.limpa_dados(numero_processo)
            tipo_instrumento = robo.limpa_dados(tipo_instrumento)

            # input para reset do arquivo JSON
            reset = input('Deseja resetar o robô? s/n: ')
            if reset.lower() == 's':
                robo.reset(arquivo_log=arquivo_log)

            # Em caso de parada o programa recomeça da última linha iterada
            progresso = robo.carrega_progresso(arquivo_log)
            # Inicia o processo de consulta do instrumento
            robo.consulta_instrumento()

            inicio_range = 0
            if progresso["indice"] > 0:
                inicio_range = progresso["indice"] + 1


            for indice in range(inicio_range, max_linha):
                eta()
                if situacional[indice] != '':
                    continue
                try:
                    # Cria pasta com número do processo
                    caminho_pasta = robo.criar_pasta(nome_pasta=numero_processo[indice],
                                                     caminho_pasta_onedrive=caminho_pasta_onedrive,
                                                     tipo_instrumento=tipo_instrumento[indice])

                    # Executa pesquisa dos termos e salva os resultados na pasta "caminho_pasta"
                    robo.loop_de_pesquisa(
                        numero_processo=numero_processo[indice],
                        caminho_pasta=caminho_pasta,
                        pasta_download=pasta_download,
                    )
                    try:
                        robo.loop_esclarecimento(pasta_download=pasta_download,
                                                 caminho_pasta=caminho_pasta,
                                                 numero_processo=numero_processo[indice]
                                                 )
                    except Exception as e:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                        print(
                            f"\n❌ Erro ao processar esclarecimento no {indice}: ({numero_processo[indice]}).\n"
                            f" {e[:50]}")

                    # Confirma se houve atualização na pasta e envia email para o técnico
                    confirma_email = list(robo.condicao_email(numero_processo=numero_processo[indice],
                                                              caminho_pasta=caminho_pasta))
                    if confirma_email:
                        robo.salva_progresso(arquivo_log, numero_processo[indice], confirma_email[2],
                                             indice=indice)

                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(f"\n❌ Erro ao processar o índice {indice} ({numero_processo[indice]}): {e}")
                    robo.consulta_instrumento()
                    continue  # Continua para o próximo processo


        start_time = time.time()

        main()

        end_time = time.time()
        tempo_total = end_time - start_time
        horas = int(tempo_total // 3600)
        minutos = int((tempo_total % 3600) // 60)
        segundos = int(tempo_total % 60)
        print(f'⏳ Tempo de execução: {horas}h {minutos}m {segundos}s')


    # busca dados de quantos arquivos são esperados em cada proposta
    elif func == 10:
        class Robo:
            def __init__(self):
                """
                Inicializa o objeto Robo, configurando e iniciando o driver do Chrome.
                """
                try:
                    # Configuração do registro
                    # self.arquivo_registro = ''
                    # Inicia as opções do Chrome
                    self.chrome_options = webdriver.ChromeOptions()
                    # Endereço de depuração para conexão com o Chrome
                    self.chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
                    # Inicializa o driver do Chrome com as opções e o gerenciador de drivers
                    self.driver = webdriver.Chrome(
                        service=Service(ChromeDriverManager().install()),
                        options=self.chrome_options)
                    self.driver.switch_to.window(self.driver.window_handles[0])

                    print("✅ Conectado ao navegador existente com sucesso.")
                except WebDriverException as e:
                    # Imprime mensagem de erro se a conexão falhar
                    print(f"❌ Erro ao conectar ao navegador existente: {e}")
                    # Define o driver como None em caso de falha na conexão
                    self.driver = None

            # Chama a função do webdriver com wait element to be clickable
            def webdriver_element_wait(self, xpath: str, num_element: int = 1):
                """
                        Espera até que um elemento web esteja clicável, usando um tempo limite máximo de 3 segundos.

                        Args:
                            xpath: O seletor XPath do elemento.
                            num_element: Identificador do npúmero de elementos que se espera serem retornados

                        Returns:
                            O elemento web clicável, ou lança uma exceção TimeoutException se o tempo limite for atingido.

                        Raises:
                            TimeoutException: Se o elemento não estiver clicável dentro do tempo limite.
                        """
                # Cria uma instância de WebDriverWait com o driver e o tempo limite e espera o elemento ser clicável
                if num_element == 1:
                    return WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable((By.XPATH, xpath)))
                else:
                    return WebDriverWait(self.driver, 3).until(
                        lambda driver: [elem for elem in driver.find_elements(By.XPATH, xpath)]
                    )

            # Navega até a página de busca do instrumento ou proposta
            def consulta_instrumento(self):
                """
                       Navega pelas abas do sistema até a página de busca de processos.

                       Esta função clica nas abas principal e secundária para acessar a página
                       onde é possível realizar a busca de processos.
                       """
                # Reseta para página inicial
                try:
                    reset = WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable(
                        (By.XPATH, '//*[@id="header"]')))
                    if reset:
                        img = reset.find_element(By.XPATH, '//*[@id="logo"]/a/img')
                        action = ActionChains(self.driver)
                        action.move_to_element(img).perform()
                        reset.find_element(By.TAG_NAME, 'a').click()
                        print(Fore.MAGENTA + "\n✅ Processo resetado com sucesso !")
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}")
                    print(Fore.RED + f'🔄❌ Falha ao resetar.\nErro: {type(e).__name__}\n{str(e)[:100]}')

                # [0] Excução; [1] Consultar Pré-Instrumento/Instrumento
                xpaths = ['//*[@id="menuPrincipal"]/div[1]/div[4]',
                          '//*[@id="contentMenu"]/div[1]/ul/li[6]/a'
                          ]
                try:
                    for idx in range(len(xpaths)):
                        self.webdriver_element_wait(xpaths[idx]).click()
                    print(f"{Fore.GREEN}✅ Sucesso em acessar a página de busca de processo{Style.RESET_ALL}")
                except NoSuchElementException:
                    pass
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(Fore.RED + f'🔴📄 Instrumento indisponível. \nErro: {e}{Style.RESET_ALL}')
                    sys.exit(1)

            def campo_pesquisa(self, numero_processo):
                try:
                    # Seleciona campo de consulta/pesquisa, insere o número de proposta/instrumento e da ENTER
                    campo_pesquisa = self.webdriver_element_wait('//*[@id="consultarNumeroConvenio"]')
                    campo_pesquisa.clear()
                    campo_pesquisa.send_keys(numero_processo)
                    campo_pesquisa.send_keys(Keys.ENTER)

                    # Acessa o item proposta/instrumento
                    acessa_item = self.webdriver_element_wait('/html/body/div[3]/div[14]/div[3]/div['
                                                              '3]/table/tbody/tr/td/div/a')
                    acessa_item.click()
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(
                        f' Falha ao inserir número de processo no campo de pesquisa. Erro: {type(e).__name__}')

            def busca_convenio(self):
                print('\n🔁🔍 Executando loop de pesquisa de convênio')
                # Seleciona aba primária, após acessar processo/instrumento. Aba Projeto Básico/Termo de referência
                termo_referencia = self.webdriver_element_wait('/html[1]/body[1]/div[3]/div[14]/'
                                                               'div[1]/div[1]/div[1]/a[4]/div[1]/span[1]/span[1]')
                termo_referencia.click()

                # Aba Projeto Básico/Termo de referência
                aba_termo_referencia = self.webdriver_element_wait('/html[1]/body[1]/div[3]/div[14]/div[1]'
                                                                   '/div[1]/div[2]/a[9]/div[1]/span[1]/span[1]')
                aba_termo_referencia.click()

            def acessa_anexos(self):
                try:
                    # Executa pesquisa de anexos
                    print('🔁📎 Executando pesquisa de anexos')

                    # Aba Plano de trabalho
                    self.webdriver_element_wait('//*[@id="div_997366806"]').click()

                    # Seleciona aba anexos
                    self.webdriver_element_wait('//*[@id="menu_link_997366806_1965609892"]/div').click()

                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(f"Ocorreu um erro ao executar ao pesquisar anexos: {type(e).__name__}"
                          f".\n Erro {e[:50]}")

            # Pesquisa o termo de fomento listado na planilha e executa download e transferência caso exista algúm.
            def loop_de_pesquisa(self, numero_processo: str):
                def conta_paginas(tabela):
                    # Diz quantas páginas tem
                    try:
                        paginas = tabela.find_element(By.TAG_NAME, 'span').text
                        paginas = paginas.split('(')[1]
                        paginas = int(paginas.split('i')[0].strip())
                        return paginas
                    except Exception:
                        print('No pages found')
                        return 0

                self.campo_pesquisa(numero_processo=numero_processo)

                paginas = 0
                try:
                    self.acessa_anexos()

                    # Seleciona lista de anexos execução e manda baixar os arquivos
                    try:

                        print('\n🔁📎 Executando pesquisa de anexos execução')

                        time.sleep(0.3)
                        botao_lista_execucao = self.webdriver_element_wait('//tbody//tr//input[2]')

                        if botao_lista_execucao.is_displayed() or botao_lista_execucao.is_enabled():
                            try:
                                # Seleciona lista de anexos execução e acessa a mesma
                                botao_lista_execucao.click()
                            except TimeoutException:
                                print(f"Timeout waiting for element: {'//tbody//tr//input[2]'}")
                            except Exception as e:  # Catch other potential exceptions
                                exc_type, exc_value, exc_tb = sys.exc_info()
                                print(
                                    Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                                print(f"🤷‍♂️❌ Erro ao tentar entar na lista de anexos execução: {e}")

                            # Encontra a tabela de anexos
                            tabela = WebDriverWait(self.driver, 10).until(
                                EC.presence_of_element_located((By.ID, 'listaAnexos')))
                            # Diz quantas páginas tem
                            paginas = conta_paginas(tabela)


                        else:
                            # Volta para a aba de consulta (começo do loop) caso não tenha lista de execução
                            self.webdriver_element_wait('/html[1]/body[1]/div[3]/div[2]/div[6]/a[2]').click()

                        print(
                            f"\n{Fore.GREEN}✅ Loop de pesquisa concluído para o processo:"
                            f" {numero_processo}{Style.RESET_ALL}\n")

                        return paginas

                    except Exception as er:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                        print(f'❌ Falha ao acessar documentos de execução.'
                              f'\nErro:{type(er).__name__}, {str(er)[:50]}')
                        self.consulta_instrumento()

                except TimeoutException as t:
                    print(f'TIMEOUT {t[:50]}')
                    self.consulta_instrumento()
                except Exception as erro:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(
                        f'❌ Falha ao acessar documentos de execução. Erro:{type(erro).__name__}\n{str(erro)[:50]}')
                    self.consulta_instrumento()

            # Limpa o nome do arquivo
            @staticmethod
            def clean_filename(filename):
                # Remove or replace invalid characters for Windows filenames
                return re.sub(r'[\\/*?:"<>|]', "_", filename)

            # Salva a tela de esclarecimento detalhado.
            def loop_esclarecimento(self, numero_processo: str):
                # Xpath [0] Acomp. e Fiscalização // [1] Esclarecimento
                loop_list = ['//*[@id="menuInterno"]/div/div[7]',
                             '//*[@id="contentMenuInterno"]/div/div[1]/ul/li[3]/a']

                conta_resposta = 0

                for j in loop_list:
                    try:
                        self.webdriver_element_wait(j).click()
                    except TimeoutException as e:
                        print(f'❌ Falha ao acessar documentos de esclarecimento. Erro:{type(e).__name__}')

                print(f'🖨️🖼️ Imprimindo tela do processo {numero_processo}.')

                try:
                    # Encontra a tabela de anexos
                    tabela = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.ID, 'esclarecimentos')))
                    # Diz quantas páginas tem
                    paginas = self.conta_paginas(tabela)

                    try:
                        for pagina in range(1, paginas + 1):
                            if pagina > 1:
                                element = WebDriverWait(self.driver, 10).until(
                                    EC.element_to_be_clickable((By.LINK_TEXT, f'{pagina}')))
                                element.click()
                                print(f'\nAcessando página {pagina}\n ')

                                # Encontra a tabela na página atual
                            tabela = WebDriverWait(self.driver, 3).until(
                                EC.presence_of_element_located((By.ID, 'esclarecimentos')))
                            # Encontra todas as linhas da tabela atual
                            linhas = tabela.find_elements(By.TAG_NAME, 'tr')

                            previous_text = ''

                            for indice in range(1, len(linhas)):
                                indice = int(indice)
                                try:
                                    # Refresh table and rows reference periodically
                                    if indice >= 1:
                                        tabela = WebDriverWait(self.driver, 3).until(
                                            EC.presence_of_element_located((By.ID, 'esclarecimentos')))
                                        linhas = tabela.find_elements(By.TAG_NAME, 'tr')

                                    linha = linhas[indice]
                                    if linha.text.split('\n')[3] == previous_text:
                                        continue

                                    botao_detalhar = linha.find_element(By.CLASS_NAME, 'buttonLink')
                                    botao_detalhar.click()
                                except Exception as error:
                                    exc_type, exc_value, exc_tb = sys.exc_info()
                                    print(
                                        Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                                    print(
                                        f"❌ Erro ao processar a linha nº{indice},"
                                        f" erro: {type(error).__name__}\n{str(error)[:80]}")
                                    break

                                try:
                                    time.sleep(0.5)
                                    # Encontra a tabela na página atual
                                    tabela_resp = self.driver.find_element(
                                        By.XPATH,
                                        '/html/body/div[3]/div[14]/div[3]/div/div/form/table/tbody/tr['
                                        '18]/td/div[1]')
                                    # Encontra todas as linhas da tabela atual
                                    linhas_resp = tabela_resp.find_elements(By.XPATH, './/tbody/tr')

                                    print(len(linhas_resp))
                                    conta_resposta += len(linhas_resp)

                                except NoSuchElementException:
                                    pass
                                except TimeoutException:
                                    pass
                                except Exception as err:
                                    exc_type, exc_value, exc_tb = sys.exc_info()
                                    print(
                                        Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}"
                                                   f"{Style.RESET_ALL}")
                                    print(Fore.RED +
                                        f'❌ Erro de download: Termo {str(err)[:100]}.'
                                        f' Err: {type(err).__name__}{Style.RESET_ALL}')

                                self.driver.execute_script(
                                    "window.scrollTo(0, document.body.scrollHeight);")
                                pyautogui.hotkey('alt', 'left')

                    except Exception as error:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                        print(f"❌ Erro ao buscar nova página de esclarecimento: {error}.\nErr:"
                              f" {type(error).__name__}")
                except Exception as error:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(f"❌ Erro ao econtrar elementos na página de esclarecimentos: {error}. Err:"
                          f" {type(error).__name__}")

                self.consulta_instrumento()

                print(
                    f"\n{Fore.GREEN}✅ Loop de esclarecimento concluído para o processo:"
                    f" {numero_processo}{Style.RESET_ALL}\n")

                return conta_resposta

            # Finds which locator to use
            def find_button_with_retry(self):
                # Define all possible locator strategies and values
                locators = [
                    (By.ID, 'form_submit'),
                    (By.NAME, 'detalharEsclarecimentoConvenioDadosDoEsclarecimentoVoltarForm'),
                    (By.XPATH, '//input[@value="Voltar"]'),
                    (By.XPATH, '//td[@class="FormLinhaBotoes"]/input'),
                    (By.CLASS_NAME, 'FormLinhaBotoes'),  # Will need additional find after
                    (By.XPATH, '//input[contains(@onclick, "setaAcao")]'),
                    (By.XPATH, '//*[@id="form_submit"]')
                ]

                for locator in locators:
                    try:
                        print(f'Botão localizado com o seletor {locator[0]}')
                        return self.driver.find_element(*locator)
                    except Exception as e:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                        print(f"Falha com localizador {locator}: {str(e)[:100]}")
                        continue

                raise NoSuchElementException("Could not find button using any locator strategy")

            # Conta quantas páginas tem para iterar sobre
            @staticmethod
            def conta_paginas(tabela):
                # Diz quantas páginas tem
                try:
                    paginas = tabela.find_element(By.TAG_NAME, 'span').text
                    paginas = paginas.split('(')[0]
                    paginas.strip()
                    paginas = int(paginas[-3:])
                    return paginas
                except NoSuchElementException:
                    paginas = 1
                    return paginas

            @staticmethod
            def converter_data(data: str) -> datetime | None:
                """
                Função auxiliar para converter uma string de data em um objeto datetime,
                tratando diferentes formatos.

                Args:
                    data: String de data a ser convertida.

                Returns:
                    Um objeto datetime correspondente à data fornecida ou None se o formato for inválido.
                """
                formatos_tentativa = [
                    '%Y-%m-%d %H:%M:%S',  # Formato com hora
                    '%Y-%m-%d',  # Formato sem hora
                    '%Y-%m-%dT%H:%M:%S.%fZ',  # Formato ISO 8601 com microssegundos e 'Z'
                    '%d/%m/%Y'  # Formato de data BR
                ]

                for formato in formatos_tentativa:
                    try:
                        return datetime.strptime(data, formato)
                    except ValueError:
                        pass  # Tenta o próximo formato

                print(f"⚠️📆 Formato de data inválido: {data}")
                return None

            @staticmethod
            def data_hoje():
                # pega a data do dia que está executando o programa
                agora = datetime.now()
                data_hoje = datetime(agora.year, agora.month, agora.day)
                return data_hoje

            @staticmethod
            def extrair_dados_excel(caminho_arquivo_fonte, busca_id: str, tipo_instrumento_id: str,
                                    situacional_id: str) -> tuple:
                """
                   Lê os contatos de uma planilha Excel e executa ações baseadas nos dados extraídos.

                   Args:
                       caminho_arquivo_fonte (str): Caminho do arquivo Excel que será lido.
                       busca_id (str): Nome da coluna que contém os números de processo.
                       tipo_instrumento_id (str): Nome da coluna que contém o tipo de processo.
                   """

                pd.set_option('future.no_silent_downcasting', True)
                dados_processo = pd.read_excel(caminho_arquivo_fonte, dtype=str)
                dados_processo = dados_processo.fillna('')
                dados_processo = dados_processo.replace({u'\xa0': ''})
                dados_processo = dados_processo.infer_objects(copy=False)

                # Cria um lista para cada coluna do arquivo xlsx
                numero_processo = list()
                tipo_instrumento = list()
                situacional = list()

                try:
                    # Itera a planilha e armazena os dados em listas
                    for indice, linha in dados_processo.iterrows():  # Assume que a primeira linha e um cabeçalho
                        numero_processo.append(linha[busca_id])  # Busca o número do processo
                        tipo_instrumento.append(linha[tipo_instrumento_id])  # Busca o tipo de instrumento
                        situacional.append(linha[situacional_id])  # Busca a situação do processo
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(f"❌ Erro de leitura encontrado, erro: {e}")

                return numero_processo, tipo_instrumento, situacional

            # Limpa os dados que vem da planilha
            def limpa_dados(self, lista: list):
                """
                Limpa os elementos de uma lista, removendo quebras de linha, espaços extras e
                substituindo valores NaN por strings vazias.

                :param lista: Lista contendo os elementos a serem processados.
                :return: Uma nova lista com os elementos limpos.
                """
                lista_limpa = []
                for i in lista:
                    if pd.isna(i):
                        lista_limpa.append('')
                        continue
                    # Remove quebra de linha "\n"
                    i_limpo = i.replace('\n', '').replace('\r', '').strip()
                    # Anexa à lista limpa
                    lista_limpa.append(i_limpo)
                return lista_limpa

            # Salva o progresso em um arquivo json
            def salva_progresso(self, processo_visitado: str, indice: int, cont_arq_anexo: int,
                                cont_escl:int, arquivo_log):
                """
                Salva o progresso atual em um arquivo JSON.

                :param arquivo_log: Endereço do arquivo JSON
                :param processo_visitado: Lista de processos concluídos.
                :param arquivos_baixados: Lista de arquivos baixados.
                :param indice: Diz qual linha o programa iterou por último.
                """
                # Carrega os dados antigos

                try:
                    dados_log = self.carrega_progresso(arquivo_log=arquivo_log)

                    if dados_log:
                        # Carrega os dados novos
                        dados_log[processo_visitado] = {
                            "arquivos_anexo": cont_arq_anexo,
                            "arquivos_esclarecimento": cont_escl,
                            "indice": indice
                        }


                        with open(arquivo_log, 'w', encoding='utf-8') as arq:
                            json.dump(dados_log, arq, indent=4, ensure_ascii=False)
                        print(f"💾 Processo salvo: {processo_visitado} (Índice: {indice})")
                        print(f"   Arquivos anexo: {cont_arq_anexo}, Esclarecimento: {cont_escl}")
                    else:
                        print(f'Falha ao carregar os arquivos. Retornado {type(dados_log)}')
                        sys.exit(0)
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(f"\n❌ Erro ao salvar progresso, tipo: {type(e).__name__}\n{str(e)[:100]}")


            # Carrega os dados do arquivo JSON que sereve como Cartão de Memória
            def carrega_progresso(self, arquivo_log: str):
                """
                    Carrega o progresso do arquivo JSON.

                    :param arquivo_log: Endereço do arquivo JSON

                    :return: Um dicionário contendo os dados de progresso.
                             Se o arquivo não existir, retorna valores padrão.
                """
                try:
                    if not os.path.exists(arquivo_log):
                        print("⚠️  Arquivo de log não existe, retornando dict vazio")
                        return {}

                    with open(arquivo_log, 'r', encoding='utf-8') as arq:
                        print('Success reading file !')
                        return json.load(arq)
                except json.JSONDecodeError as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")
                    print(f"JSON decode error: {e}")
                    print(f"Content that failed to parse: {type(e).__name__}\n{str(e)[:100]}")
                    return None

                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")
                    print(f"Other error tipo: {type(e).__name__}\n{str(e)[:100]}")
                    return None


            # Reseta o arquivo JSON
            def reset(self, arquivo_log: str):
                """
                Reseta o progresso atual em um arquivo JSON.

                :param arquivo_log: Endereço do arquivo JSON

                """
                dados_vazios = {
                    "processo_visitado": [],
                    "indice": 0
                }
                try:
                    # Salva os dados vazios no arquivo JSON
                    with open(arquivo_log, 'w', encoding='utf-8') as arq:
                        json.dump(dados_vazios, arq, indent=4)
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")
                    print(f"Other error tipo: {type(e).__name__}\n{str(e)[:100]}")
                    return None
            # Verifica as condições para mandar um e-mail para o técnico
            def condicao_email(self, numero_processo: str, caminho_pasta: str):
                """
                Verifica se há arquivos modificados na data de hoje dentro de um diretório específico.

                :param numero_processo: Número do processo relacionado aos arquivos.
                :param caminho_pasta: Caminho da pasta onde os arquivos estão armazenados.
                :return: Uma tupla contendo (numero_processo, caminho_pasta, lista de arquivos modificados hoje).
                         Caso não haja arquivos modificados hoje, retorna uma lista vazia.
                """

                # lista que guarda os arquivos novos, na função ENVIAR_EMAIL_TECNICO recebe o nome de lista_documentos
                docs_atuais = []
                try:
                    # Data de hoje
                    hoje = self.data_hoje()
                    # Itera os arquivos da pasta para buscar a data de modificação individual
                    for arq_nome in os.listdir(caminho_pasta):
                        arq_caminho = os.path.join(caminho_pasta, arq_nome)
                        # Pula diretórios
                        if os.path.isfile(arq_caminho):
                            data_mod = datetime.fromtimestamp(os.path.getmtime(arq_caminho))
                            # Compara as datas de modificação dos arquivos
                            if data_mod >= hoje:
                                docs_atuais.append(arq_nome)
                    if docs_atuais:
                        print(f"📂✨ Documentos novos encontrados para o processo {numero_processo}")
                        return numero_processo, caminho_pasta, docs_atuais

                    else:

                        print(f"⚠️Nenhum documento novo encontrado para o processo {numero_processo}.\n")
                        return []
                except Exception:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(f"Error occurred at line: {exc_tb.tb_lineno}")
                    print(f"⚠️Nenhum documento novo encontrado para o processo {numero_processo}.\n")
                    return []


        def main() -> None:
            # Caminho do arquivo .xlsx que contem os dados necessários para rodar o robô
            caminho_arquivo_fonte = (
                r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\Automações SNEAELIS\Analise_Custos_Exec_Print\source\RTMA Passivo 2024 - '
                r'PROJETOS E PROGRAMAS (3) Atualizado em Julho de.2025.xlsx ')

            arquivo_log = (
                r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\Automações SNEAELIS\Analise_Custos_Exec_Print\source\verify_log.json')

            try:
                # Instancia um objeto da classe Robo
                robo = Robo()
                # Extrai dados de colunas específicas do Excel
            except Exception as e:
                exc_type, exc_value, exc_tb = sys.exc_info()
                print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                print(f"\n‼️ Erro fatal ao iniciar o robô: {e}")
                sys.exit("Parando o programa.")

            numero_processo, tipo_instrumento, situacional = robo.extrair_dados_excel(
                caminho_arquivo_fonte=caminho_arquivo_fonte,
                busca_id='Instrumento nº',
                tipo_instrumento_id='Regime Jurídico do Instrumento (Modalidade)',
                situacional_id='SITUACIONAL'
            )

            max_linha = len(numero_processo)

            # Pós-processamento dos dados para não haver erros na execução do programa
            numero_processo = robo.limpa_dados(numero_processo)

            # Inicia o processo de consulta do instrumento
            robo.consulta_instrumento()

            robo.reset(arquivo_log=arquivo_log)

            inicio_range = 0

            for indice in range(inicio_range, max_linha):
                if situacional[indice] != '':
                    continue
                try:
                    # Executa pesquisa dos termos e salva os resultados na pasta "caminho_pasta"
                    pagina = robo.loop_de_pesquisa(numero_processo=numero_processo[indice])

                    conta_resposta = robo.loop_esclarecimento(numero_processo=numero_processo[indice])

                    robo.salva_progresso(indice=indice,
                                         processo_visitado=numero_processo[indice],
                                         cont_arq_anexo=pagina,
                                         cont_escl=conta_resposta,
                                         arquivo_log=arquivo_log
                                         )

                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(Fore.RED + f"Error occurred at line: {exc_tb.tb_lineno}{Style.RESET_ALL}")

                    print(f"\n❌ Erro ao processar o índice {indice} ({numero_processo[indice]}): {e}")
                    robo.consulta_instrumento()
                    continue  # Continua para o próximo processo

        main()


    # Check files on analise_custos_print
    elif func == 11:
        def conta_arquivo_zip(zip_path: str) -> int:
            try:
                with zipfile.ZipFile(zip_path, 'r') as zf:
                    return  len([f for f in zf.namelist() if not f.endswith('/')])
            except Exception:
                print(f'No files foun at {os.path.basename(zip_path)}')
                return 0

        def conta_arquivo_esclarecimento(dir_path: str) -> tuple:
            pattern = re.compile(r'\d{2}_\d{2}_\d{4}_\d{1,2}_\d{4}')
            counter_zip = 0
            counter_esc = 0

            for file in os.listdir(dir_path):
                if file.lower().endswith(".zip") and file.split('.')[0] == os.path.basename(dir_path):
                   counter_zip = conta_arquivo_zip(os.path.join(dir_path, file))
                   continue
                elif pattern.search(file):
                    continue

                counter_esc += 1

            return counter_esc, counter_zip

        def compara_com_json(dir_path: str, json_path: str, json_out_path: str):
            key = os.path.basename(dir_path).split('_')[0]

            with open(json_path, 'r', encoding='utf-8') as f:
                dados = json.load(f)

            if key not in dados:
                print(f"⚠️ Chave {key} não encontrada no JSON.")
                return

            arq_esclarecimento, arq_anexo = conta_arquivo_esclarecimento(dir_path)

            esperado_anexo = dados[key].get('arquivos_anexo', 0)
            esperado_esclarecimento = dados[key].get('arquivos_esclarecimento', 0)

            # Comparação
            has_diff = False

            if arq_anexo != esperado_anexo or arq_esclarecimento != esperado_esclarecimento:
                print(f"📂 Processo {key}")
                print(f"   Anexo → encontrado: {arq_anexo}, esperado: {esperado_anexo}")
                print(
                    f"   Esclarecimento → encontrado: {arq_esclarecimento},"
                    f" esperado: {esperado_esclarecimento}")
                has_diff = True

            else:
                print(f"✅ Processo {key}: sem diferenças")

            resultado = {
                "processo": key,
                "encontrado": {
                    "arquivos_anexo": arq_anexo,
                    "arquivos_esclarecimento": arq_esclarecimento,
                },
                "esperado": {
                    "arquivos_anexo": esperado_anexo,
                    "arquivos_esclarecimento": esperado_esclarecimento,
                }
            }
            if os.path.exists(json_out_path):
                with open(json_out_path, 'r', encoding='utf-8') as f_load:
                    try:
                        dados = json.load(f_load)
                    except json.JSONDecodeError:
                        dados = []
            else:
                dados = []

            if not isinstance(dados, list):
                dados = [dados]

            dados.append(resultado)

            with open(json_out_path, 'w', encoding='utf-8') as f_out:
                json.dump(dados, f_out, indent=4, ensure_ascii=False)


            return has_diff


        def main():
            json_out_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\Automações '
             r'SNEAELIS\Analise_Custos_Exec_Print\source\comparison_log.json')

            dir_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\Automações '
             r'SNEAELIS\Analise_Custos_Exec_Print')

            json_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\Automações '
             r'SNEAELIS\Analise_Custos_Exec_Print\source\verify_log.json')

            xlsx_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\Automações '
             r'SNEAELIS\Analise_Custos_Exec_Print\source\RTMA Passivo 2024 - PROJETOS E PROGRAMAS (3) '
                         r'Atualizado em Julho de.2025.xlsx')

            count_diff = 0
            to_redo_list = []

            df = pd.read_excel(xlsx_path, dtype=str)
            isnt_num = list()
            missing_instruments = list()

            for indice, linha in df.iterrows():  # Assume que a primeira linha e um cabeçalho
                if (linha['SITUACIONAL']) != '':
                    isnt_num.append(linha['Instrumento nº'])


            for root, _, files in os.walk(dir_path):
                for d in _:
                    if compara_com_json(dir_path=os.path.join(root, d),
                                     json_path=json_path,
                                     json_out_path=json_out_path):
                        count_diff += 1
                        to_redo_list.append(d)
                    if d not in isnt_num:
                        missing_instruments.append(d)

            print(f"\n🔎 Total de diretórios com diferenças: {count_diff}\n\nList to redo:"
                  f"{to_redo_list}\n\nDid not iterate: {missing_instruments}, {len(missing_instruments)}")


        main()


    # Testing logger funtion
    elif func == 12:
        class TestRoboPAD:
            def __init__(self):
                self.logger = self.setup_logger()
                self.logger.info("🔧 TestRoboPAD instance created")

            def setup_logger(self, level=logging.INFO):
                log_file_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                                 r'Social\SNEAELIS - Robô PAD')

                # Create logs directory if it doesn't exist
                log_file_name = f'log_PAD_{datetime.now().strftime('%d_%m_%Y')}.log'

                # Sends to specific directory
                log_file = os.path.join(log_file_path, log_file_name)

                if not os.path.exists(log_file_path):
                    os.makedirs(log_file_path)
                    print(f"✅ Directory created/verified: {log_file_path}")

                logger = logging.getLogger()
                logger.setLevel(level)

                formatter = logging.Formatter(
                    '%(asctime)s | | %(message)s',
                    datefmt='%Y-%m-%d  %H:%m'
                )
                og_format = formatter.format
                formatter.format = lambda record: og_format(record) + '\n' + '─' * 100

                # File handler with rotation
                file_handler = logging.handlers.RotatingFileHandler(
                    log_file,
                    maxBytes=10485760,
                    backupCount=5,
                    encoding='utf-8'
                )
                file_handler.setFormatter(formatter)

                console_handler = logging.StreamHandler()
                console_handler.setFormatter(formatter)

                logger.addHandler(file_handler)
                logger.addHandler(console_handler)

                if os.path.exists(log_file):
                    print(f"🎉 SUCCESS! Log file created at: {log_file}")
                    print(f"📊 File size: {os.path.getsize(log_file)} bytes")
                    return logger
                else:
                    print(f"❌ File not created at: {log_file}")

            def simulate_successful_process(self, process_id):
                """Simulate a successful process"""
                self.logger.info(f"🚀 Starting successful process: {process_id}")

                # Simulate steps
                self.logger.info("📋 Step 1: Validating input data")
                time.sleep(0.1)  # Simulate work

                self.logger.info("🔍 Step 2: Searching for buttons")
                time.sleep(0.1)

                self.logger.info("✅ Step 3: Process completed successfully")
                self.logger.info(f"🎉 Process {process_id} finished")

                return True

            def test_unicode_characters(self):
                """Test that special characters work"""
                self.logger.info("🔤 Testing Unicode characters:")
                self.logger.info("✅ Normal text")
                self.logger.info("🔘 Button emoji test")
                self.logger.info("💤 Sleep emoji test")
                self.logger.info("🚀 Rocket emoji test")
                self.logger.info("⚠️  Warning emoji test")
                self.logger.info("💥 Error emoji test")
                self.logger.info("🎉 Success emoji test")

        test_ = TestRoboPAD()
        test_.simulate_successful_process(3126516)
        test_.test_unicode_characters()


    # Create directories
    elif func == 13:
        # CONFIGURATION
            # Path to your Excel file
        xl_file_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                        r'Social\Teste001\Processos Gerar Parecer TF.xlsm')
            #  Column name for process number
        precess_col = 'Processo'
            #  Column name for xxxxxx-xxxx
        proposal_col = 'Nº Proposta'
            # Where to create the new directories
        output_dir_parent = (
            fr'C:\Users\felipe.rsouza\Documents\fabi')
            # The directory whose contents are to be copied
        std_dir_path = [
            # To copy folder when "Propostas TF"
            r'C:\Users\felipe.rsouza\Documents\Convênio\29378-2025 - 71000062161202500\Portarias e Parecer '
            r'Referencial CONJUR',
            # To copy folder when "Propostas Convênio"
            r'C:\Users\felipe.rsouza\Documents\Convênio\29378-2025 - 71000062161202500\Portarias e Parecer '
            r'Referencial CONJUR'
        ]
        # SUBDIRECTORIES TO CREATE
        subdirs = [
            'Planilha de Custos e Cotações',
            'Portarias e Parecer Referencial CONJUR',
            'Requisitos de Celebração'
        ]
        sheet_names = ['Propostas TF', 'Propostas Convênio']
        dir_name_sufix = ['Termo de Fomento', 'Convênio']

        for i, sheet in enumerate(sheet_names):
            output_dir_path = os.path.join(output_dir_parent, dir_name_sufix[i])
            df = pd.read_excel(xl_file_path, dtype=str, sheet_name=sheet)
            replace_ = ['.', '/', '-']

            for idx, row in df.iterrows():
                process = str(row[precess_col]).strip()
                for _ in replace_:
                    process = process.replace(_, '')
                proposal = str(row[proposal_col]).strip().replace('/','_')
                main_dir = f'{proposal} - {process}'
                main_dir_path = os.path.join(output_dir_path, main_dir)
                os.makedirs(main_dir_path, exist_ok=True)

                for subdir in subdirs:
                    subdir_path = os.path.join(main_dir_path, subdir)
                    if subdir == 'Portarias e Parecer Referencial CONJUR':
                        # Copy the entire source directory structure and files
                        os.makedirs(subdir_path, exist_ok=True)
                        shutil.copytree(std_dir_path[i], subdir_path, dirs_exist_ok=True)

                    else:
                        os.makedirs(subdir_path, exist_ok=True)

        print("Directories created successfully.")


    if func == 14:
        def compress_pdf_fitz(input_path, output_path, dpi=150, quality=80):
            """
            Compress PDF using PyMuPDF - no poppler required
            """
            # Open the original PDF
            doc = fitz.open(input_path)

            # Create a new PDF for compressed output
            writer = fitz.open()

            for page_num in range(len(doc)):
                page = doc[page_num]

                # Render page as image with specified DPI
                mat = fitz.Matrix(dpi / 72, dpi / 72)  # Convert DPI to matrix
                pix = page.get_pixmap(matrix=mat)

                # Convert to bytes and then to PIL Image for compression
                img_data = pix.tobytes("ppm")
                pil_img = Image.open(io.BytesIO(img_data))

                # Convert to RGB if necessary
                if pil_img.mode != 'RGB':
                    pil_img = pil_img.convert('RGB')

                # Compress image
                output_buffer = io.BytesIO()
                pil_img.save(output_buffer, format='JPEG', quality=quality, optimize=True)
                compressed_img_data = output_buffer.getvalue()

                # Create new page with compressed image
                new_page = writer.new_page(width=page.rect.width, height=page.rect.height)
                new_page.insert_image(new_page.rect, stream=compressed_img_data)

            # Save with compression options
            writer.save(output_path, garbage=4, deflate=True, clean=True)

            doc.close()
            writer.close()

            # Print compression results
            original_size = os.path.getsize(input_path)
            compressed_size = os.path.getsize(output_path)
            reduction = (1 - compressed_size / original_size) * 100

            print(f"Original: {original_size / 1024 / 1024:.2f} MB")
            print(f"Compressed: {compressed_size / 1024 / 1024:.2f} MB")
            print(f"Reduction: {reduction:.1f}%")


        input_path = r'C:\Users\felipe.rsouza\Documents\log_PAD\atestado nova resende 1 pag - Copia.pdf'

        output_path = r'C:\Users\felipe.rsouza\Documents\log_PAD\atestado nova resende 1 pag - resized.pdf'

        compress_pdf_fitz(input_path,
                            output_path,
                            dpi=120,
                            quality=80)
