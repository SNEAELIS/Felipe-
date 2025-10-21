import os.path
import shutil
import sys
import re
import unicodedata
import logging.handlers
import logging

import pandas as pd
import numpy as np

from datetime import datetime

from pandas import ExcelWriter

from colorama import  Fore

from thefuzz import process, fuzz

import pandas as pd

from datetime import datetime

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError
from playwright.sync_api import sync_playwright


class BreakInnerLoop(Exception):
    pass

class PWRobo:
    def __init__(self, cod_natureza_despesa:dict, cdp_url:str ="http://localhost:9222"):
        # Standard text for agreement proposal
        self.cod_natur_desp = cod_natureza_despesa

        # Connect to existing browser via Chrome DevTools Protocol (CDP)
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.connect_over_cdp(cdp_url)
        # Get all browser contexts (browser windows/profiles)
        self.context = self.browser.contexts[0] if self.browser.contexts else self.browser.new_context()
        # Get all open pages (tabs)
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        #self.block_rss()

        # Defines Logger
        self.logger = self.setup_logger()

        print(f"✅ Connected to existing Chrome instance via Playwright. Connected to page: {self.page.url}")

    @staticmethod
    def setup_logger(level=logging.INFO):
        log_file_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                         r'Social\Teste001\fabi_DFP')

        # Create logs directory if it doesn't exist
        log_file_name = f'log_Pareceres_email{datetime.now().strftime('%d_%m_%Y')}.log'

        # Sends to specific directory
        log_file = os.path.join(log_file_path, log_file_name)

        if not os.path.exists(log_file_path):
            os.makedirs(log_file_path)
            print(f"✅ Directory created/verified: {log_file_path}")

        logger = logging.getLogger()
        if logger.handlers:
            return logger

        # Avoid adding handlers multiple times
        logger.setLevel(level)

        formatter = logging.Formatter(
            '%(asctime)s | | %(message)s\n' + '─' * 100,
            datefmt='%Y-%m-%d  %H:%M'
        )

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


    def consulta_proposta(self):
        """Navigates through the system tabs to the process search page."""
        # Reseta para página inicial
        try:
            self.page.locator('xpath=//*[@id="imgLogo"]').click(timeout=8000)
        except PlaywrightTimeoutError:
             print('Already on the initial page of transferegov discricionárias.')
        except PlaywrightError as e:
            print(Fore.RED + f'🔄❌ Failed to reset.\nError: {type(e).__name__}\n{str(e)[:50]}')

        # [0] Excução; [1] Consultar Proposta
        xpaths = ['xpath=//*[@id="menuPrincipal"]/div[1]/div[3]',
                  'xpath=//*[@id="contentMenu"]/div[1]/ul/li[2]/a'
                  ]
        try:
            for xpath in xpaths:
                self.page.locator(xpath).click(timeout=8000)

        except PlaywrightError as e:
            print(Fore.RED + f'🔴📄 Instrument unavailable. \nError: {e}')
            sys.exit(1)


    def block_rss(self):
        def route_handler(route, request):
            if request.resource_type in ["image", "stylesheet", "font"]:
                route.abort()
            else:
                route.continue_()
        self.page.route('**/*', route_handler)


    def campo_pesquisa(self, numero_processo):
        try:
            # Select search field, insert proposal/instrument number and press ENTER
            campo_pesquisa_locator = self.page.locator('xpath=//*[@id="consultarNumeroProposta"]')
            campo_pesquisa_locator.fill(numero_processo)
            campo_pesquisa_locator.press('Enter')

            try:
                # Access the proposal/instrument item
                acessa_item_locator = self.page.locator('xpath=//*[@id="tbodyrow"]/tr/td[1]/div/a')
                acessa_item_locator.click(timeout=8000)
            except PlaywrightTimeoutError:
                print(f' Process number: {numero_processo}, not found.')
                self.logger.info(f'Process number: {numero_processo}, not found')
                raise BreakInnerLoop
            except PlaywrightError as e:
                print(f' Process number: {numero_processo}, not found. Error: {type(e).__name__}')
                self.logger.info(f'Process number: {numero_processo}, not found')
                raise BreakInnerLoop
        except PlaywrightError as e:
            print(f' Failed to insert process number in the search field. Error: {type(e).__name__}')


    def busca_endereco(self, cnpj_xlsx: str, num_prop: str):
        cod_municipio_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e '
                              r'Assistência Social\Teste001\municipios.xlsx')
        self.page.wait_for_timeout(500)
        try:
            print(f'🔍 Starting address search'.center(50, '-'), '\n')

            # Participants Tab
            self.page.locator('r/html/body/div[3]/div[14]/div[1]/div/div[2]/a[3]/div/span').click(
                timeout=8000)
            # Detail button
            self.page.wait_for_timeout(500)
            self.page.locator('xpath=//*[@id="form_submit"]').click(timeout=8000)

            cnjp_web = self.page.locator('xpath=//*[@id="txtCNPJ"]').text_content()
            if cnpj_xlsx != cnjp_web:
                self.logger.info(f'Proposal number: {num_prop} has CNPJ {cnjp_web}, incompatible with the spreadsheet CNPJ:'
                                 f' {cnpj_xlsx}')
                raise ValueError("Incompatible CNPJ between site and spreadsheet")

            endereco = self.page.locator('xpath=//*[@id="txtEndereco"]').text_content()
            cep = endereco.split('CEP:')[-1].replace('-','')

            # Find municipality code
            cod_municipio = ''
            pattern = r"\.\s*([^.]+?)\s*-\s*([A-Z]{2})\."
            match = re.search(pattern, endereco)
            if match:
                nome_estado = str(match.group(1)).lower()
                sigla_estado = match.group(2)
                df_mun = pd.read_excel(cod_municipio_path, dtype=str)
                linha_da_planilha = df_mun[(df_mun["MUNICÍPIO - TOM"].str.lower() ==
                                            nome_estado) & (df_mun["UF"] == sigla_estado)]
                if not linha_da_planilha.empty:
                    cod_municipio = linha_da_planilha.iloc[0, 0]  # first matching row, column A (index 0)
                else:
                    print("No match found!")

            return endereco, cep, cod_municipio
        except PlaywrightTimeoutError:
            raise BreakInnerLoop
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Error searching address:{type(e).__name__}.\n Error {str(e)[:100]}")


    def cod_mun(self, cod_municipio):
        try:
            # Wait for the new window to open
            popup_page = self.page.wait_for_event("popup")

            # Switch to the new window
            popup_page.bring_to_front()

            # Field to insert the code number
            campo_prenchimento_locator = popup_page.locator('xpath=//*[@id="consultarCodMunicipio"]')
            campo_prenchimento_locator.fill(cod_municipio)
            campo_prenchimento_locator.press('Enter')

            # Consult button
            popup_page.locator('xpath=//*[@id="form_submit"]').click(timeout=8000)
            (popup_page.locator('/html/body/div[3]/div[3]/div[2]/table/tbody/tr/td[4]/nobr/a')
                   .click(timeout=8000))

            # Close the popup window (optional, depends on desired behavior)
            popup_page.close()

            # Return to the original window
            self.page.bring_to_front()
        except PlaywrightError as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Failed to register municipality code"
                  f" {type(e).__name__}.\n Error: {str(e)[:80]}")


    def nav_plano_act_det(self):
        self.page.wait_for_timeout(1000)
        try:
            print(f'🚢 Navigating to the detailed action plan:'.center( 50, '-'), '\n')
            # Return to proposal data
            self.page.locator('xpath=//*[@id="lnkConsultaAnterior"]').click(timeout=8000)

            # Work Plan Tab
            self.page.locator('xpath=//*[@id="div_997366806"]').click(timeout=8000)

            # Select Detailed Application Plan
            self.page.locator('xpath=//*[@id="menu_link_997366806_836661414"]/div').click(timeout=8000)

        except PlaywrightError as e:
            print(f"❌ An error occurred while searching for Detailed Action Plan: {type(e).__name__}"
                  f".\n Error {str(e)[:50]}")


    # Loop para adicionar PAD da proposta
    def loop_de_pesquisa(self, df, numero_processo: str, tipo_desp: str,
                         cod_natur_desp: str, cnpj_xlsx: str, caminho_arquivo_fonte: str):
        def sanitize_txt(txt: str) -> str:
            """
               Sanitizes text by replacing special characters with form-friendly alternatives.

               Args:
                   text (str): The input text to sanitize.

               Returns:
                   str: The sanitized text.
               """
            try:
                # Step 1: Remove non-printable characters (control characters)
                txt = re.sub(r'[\x00-\x1F\x7F]', '', txt)
                # Step 2: Replace double quotes with single quotes
                txt = re.sub(r'"', "'", txt)
                # Step 3: Replace semicolons with commas
                txt = re.sub(r':', ',', txt)
                # Step 4: Normalize whitespace (replace multiple spaces with single space,
                # strip leading/trailing)
                txt = re.sub(r'\s+', ' ', txt).strip()
                # Step 5: Replace all special characters (punctuation and symbols) with spaces
                txt = re.sub(r'[^\w\s]', ' ', txt)

                return txt

            except Exception as err:
                print(f"Error sanitizing text: {str(err)[:100]}")
                return txt

        def mark_as_done(status_df, index, file_path):
            """Safely mark row as done and save to Excel"""
            try:
                index = str(index)

                status_df.loc[status_df['Index'] == index, 'Status'] = 'feito'

                # Verify the modification
                modified_value = status_df.loc[status_df['Index'] == index, 'Status'].values
                if len(modified_value) > 0 and modified_value[0] == 'feito':
                    print("✅ DataFrame modification successful")
                else:
                    print("❌ DataFrame modification failed!")

                with ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    # Save to Excel
                    status_df.to_excel(writer, index=False, sheet_name='Status')

                return True

            except Exception as err:
                print(f"❌ Error in mark_as_done: {str(err)[:100]}\n")
                return False

        def clear_digits(txt):
            """Remove all integers from text (both leading and trailing)"""
            return re.sub(r'^\d+\s*|\s*\d+$', '', txt).strip()

        # Start the instrument consultation process
        try:
            status_df = pd.read_excel(caminho_arquivo_fonte, dtype=str, sheet_name='Status')

            # Search for the process
            self.campo_pesquisa(numero_processo=numero_processo)

            # Get location data
            endereco, cep, cod_municipio = self.busca_endereco(cnpj_xlsx=cnpj_xlsx,
                                                                     num_prop=numero_processo)

            # Execute attachment search
            self.nav_plano_act_det()

            # Select execution attachment list and download files
            print('🌐 Accessing PAD filling page'.center(50, '-'), '\n')
            self.page.wait_for_timeout(1500)
            tipo_serv = self.map_tipos(tipo_desp)

            # Locate the dropdown element and select by value
            dropdown_locator = self.page.locator("#incluirBemTipoDespesa")
            dropdown_locator.select_option(tipo_serv)

            try:
                # Click to include
                btn_incluir_locator = self.page.locator('xpath=//*[@id="form_submit"]')
                btn_value =  btn_incluir_locator.get_attribute("value")
                if btn_value == 'Incluir':
                     btn_incluir_locator.click(timeout=8000)
                else:
                    self.consulta_proposta()
                    raise BreakInnerLoop
            except PlaywrightError:
                self.logger.info(f'Include button not found, skipping process: {numero_processo}')
                print(f'🔘)💤 Include button not found, skipping process: {numero_processo}')
                self.consulta_proposta()
                raise BreakInnerLoop

            lista_campos = [
                # [0] Item Description
                'xpath=//*[@id="incluirBensDescricaoItem"]',
                # [1] Expense Nature Code
                'xpath=//*[@id="incluirBensCodigoNaturezaDespesa"]',
                # [2] Supply Unit
                'xpath=//*[@id="incluirBensCodUnidadeFornecimento"]',
                # [3] Total Value
                'xpath=//*[@id="incluirBensValor"]',
                # [4] Quantity
                'xpath=//*[@id="incluirBensQuantidade"]',
                # [5] Location Address
                'xpath=//*[@id="incluirBensEndereco"]',
                # [6] CEP
                'xpath=//*[@id="incluirBensCEP"]',
            ]

            print('📝 Filling PAD'.center(50, '-'), '\n')

            for idx, row in df.iterrows():
                try:
                    alerta_valor_locator = self.page.locator('xpath=//*[@id="messages"]/div')
                    if  alerta_valor_locator.is_visible():
                        print('⚠️💸 Proposal total value exceeded the financial counterpart')
                        self.logger.info(f'Proposal total value: {numero_processo} exceeded the '
                                         f'financial counterpart')
                        self.consulta_proposta()
                        raise BreakInnerLoop
                except PlaywrightTimeoutError:
                    pass  # Element not found within timeout, continue

                # Read the guide spreadsheet, with marking of items already done
                # Check if any line has already been executed
                if status_df.iloc[idx, 1] == 'feito':
                    print(f'Line {idx} already executed. Skipping line\n')
                    continue
                print(f"Filling item nº:{idx} {str(row.iloc[2])}\n".center(50))
                try:
                    # Item Description
                    desc_item_txt = str(row.iloc[2]) + ': ' + str(row[3])
                    desc_item_txt = sanitize_txt(desc_item_txt)
                    desc_item_locator = self.page.locator(lista_campos[0])
                    desc_item_locator.fill(desc_item_txt)

                    # Expense Nature Code
                    cod_natur_locator = self.page.locator(lista_campos[1])
                    cod_natur_locator.fill(cod_natur_desp)

                    # Supply Unit
                    un_fornecimento = str(row.iloc[10]).strip().lower()
                    un_fornecimento = clear_digits(un_fornecimento)

                    if ' ' in un_fornecimento:
                        un_fornecimento = un_fornecimento.split(' ')[-1]
                    if un_fornecimento in ['mensal', 'mês', 'meses']:
                        un_fornecimento = 'MÊS'
                    elif un_fornecimento in ['unidade', 'unidades']:
                        un_fornecimento = 'UN'
                    elif un_fornecimento in ['diaria', 'diária', 'diárias']:
                        un_fornecimento = 'DIA'
                    elif un_fornecimento in ['metro']:
                        un_fornecimento = 'M'
                    elif pd.isna(un_fornecimento) or un_fornecimento == 'nan' or un_fornecimento == 'N/A':
                        print('Value not found or equal to zero\n')
                        continue

                    un_forn_field_locator = self.page.locator(lista_campos[2])
                    un_forn_field_locator.fill(un_fornecimento)

                    # Total Value
                    valor_total = (row.iloc[24])
                    if pd.isna(valor_total) or valor_total == 'nan' or valor_total == 'N/A':
                        print('Value not found or equal to zero\n')
                        continue
                    valor_total_rd = round(float(valor_total), 2)
                    formated_value = f'{valor_total_rd:.2f}'

                    valor_total_field_locator = self.page.locator(lista_campos[3])
                    valor_total_field_locator.fill(formated_value)

                    # Quantity
                    qtd = str(row.iloc[9])
                    if pd.isna(valor_total) or valor_total == 'nan' or valor_total == 'N/A':
                        print('Value not found or equal to zero\n')
                        continue
                    qtd = qtd + "00"
                    qtd_field_locator = self.page.locator(lista_campos[4])
                    qtd_field_locator.fill(qtd)

                    # Location Address
                    end_loc_locator = self.page.locator(lista_campos[5])
                    end_loc_locator.fill(endereco)

                    # CEP
                    cep_element_locator = self.page.locator(lista_campos[6])
                    cep_element_locator.fill(cep.strip())

                    # Municipality Code
                    # Assuming the button to open the municipality selection window is the third button with class "btnBusca"
                    self.page.locator("xpath=.btnBusca").nth(2).click(timeout=8000)
                    self.cod_mun(cod_municipio)

                    # "Include" button
                    self.page.locator("xpath=input#form_submit").nth(0).click(timeout=8000)
                    self.page.wait_for_timeout(2000)

                    try:
                        # Try to find the CAPTCHA error
                        error_captcha_locator = self.page.locator('.errors')
                        error_captcha_locator.wait_for(state='visible', timeout=1200)
                        # If we get here, CAPTCHA error WAS found
                        print("CAPTCHA error detected - not marking as done")
                        sys.exit("Error filling PAD.\nEnding execution")
                    except PlaywrightTimeoutError:
                        try:
                            if mark_as_done(status_df, idx, caminho_arquivo_fonte):
                                print("✅ Successfully marked and saved\n")
                            else:
                                print("❌ Failed to save\n")
                        except Exception as e:
                            print(f"❌ Error saving to Excel: {type(e).__name__}\n Erro {str(e)[:100]}\n")
                    except Exception as er:
                        print(f"❌ Failure {type(er).__name__}.\nError{er}\n")
                except PlaywrightError as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(f"Error occurred at line: {exc_tb.tb_lineno}")
                    print(f"❌ Failed to register PAD {type(e).__name__}.\n Error: {str(e)[:80]}\n")
                    return False

            # "End" button
            self.page.wait_for_timeout(1000)
            self.page.locator("input#form_submit").nth(1).click(timeout=8000)
            self.page.wait_for_timeout(1000)
            self.consulta_proposta()
            return True

        except PlaywrightTimeoutError as t:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f'TIMEOUT {str(t)[:50]}')
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            self.consulta_proposta()
        except Exception as erro:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f'❌ Failed to try to include documents. Error:{type(erro).__name__}\n'
                  f'{str(erro)[:100]}...')
            self.consulta_proposta()
            raise BreakInnerLoop


    def error_message(self):
        try:
            self.page.wait_for_selector('//*[@id="popUpLayer2"]', timeout=2000)
            close_button = self.page.locator('#popUpLayer2 img[src*="close.gif"]')
            close_button.wait_for(state="visible", timeout=5000)
            close_button.click()
            return True

        except PlaywrightTimeoutError:
            print(f"Mensagem de erro não encontrada")
            return False
        except Exception as e:
            print(f"🚨🚨 An unexpected error occurred with error message: {str(e)[:100]}\nErro name"
                  f":{type(e).__name__}")
            return False


    def map_cod_natur_desp(self, dict_cod: dict, cod: str, threshhold: int = 80) -> str:
        norm = self.normaliza_text(cod)
        choices = dict_cod.keys()

        try:
            # Check if the string, after normilized, is compatible with any key
            if norm in dict_cod:
                return dict_cod.get(norm)

            # Check if the string has any similarity with some key, threshhold is 80%
            best, score = process.extractOne(query=norm, choices=choices, scorer=fuzz.ratio)
            if score < threshhold:
                print(f"\n‼️ Failed to identify expense nature code.")
                sys.exit("Stopping the program.")

            return dict_cod.get(best)

        except Exception as e:
            print(f"\n‼️ Fatal error inserting PAD: {type(e).__name__}\nError == {str(e)[:100]}")
            sys.exit("Stopping the program.")


    @staticmethod
    def extrair_dados_excel(caminho_arquivo_fonte):
        try:
            data_frame = pd.read_excel(caminho_arquivo_fonte, dtype=str, header=None, sheet_name=0)

            return data_frame
        except Exception as e:
            print(f"🤷‍♂️❌ Error reading the excel file: {os.path.basename(caminho_arquivo_fonte)}.\n"
                  f"Error name: {type(e).__name__}\nError: {str(e)[:100]}")


    # Corrige o número da proposta que vem na planilha
    @staticmethod
    def fix_prop_num(numero_proposta):
        if '_' in numero_proposta:
            numero_proposta_fixed = numero_proposta.replace('_', '/')
        else:
            numero_proposta_fixed = numero_proposta
        return numero_proposta_fixed


    @staticmethod
    def normaliza_text(txt: str) -> str:
        if txt is None:
            return ''

        normal_text = ''.join(char for char in unicodedata.normalize('NFKD', txt) if not
        unicodedata.combining(char))

        normal_text = re.sub(r'[^a-z0-9]+', '_', normal_text)
        normal_text = re.sub(r'_+', '_', normal_text).strip()

        return normal_text


    @staticmethod
    def delete_path(path: str):
        """
        Deletes a file or directory.
        - If it's a file → delete the file.
        - If it's a directory → delete the entire directory tree.
        """
        if not os.path.exists(path):
            print(f"⚠️ Path not found: {path}")
            return

        if os.path.isfile(path):
            os.remove(path)

            print(f"🗑️ File deleted: {path}")
        elif os.path.isdir(path):
            shutil.rmtree(path)
            print(f"🗑️ Directory deleted: {path}")
        else:
            print(f"⚠️ Unknown type (not file/dir): {path}")


    # Mapeia os tipos de despesas nos quatro tipos disponíveis no transferegov
    @staticmethod
    def map_tipos(tipo):
        # Helper function with nested try-except
        def map_tipos_helper(txt: str, cat: dict) -> str:
            try:
                choices = []
                # Flatten the categories into choices with their corresponding keys
                for key, values in cat.items():
                    for value in values:
                        choices.append((value, key))

                # Extract just the text values for fuzzy matching
                choices_txt = [choice[0] for choice in choices]

                # First try exact match
                for value, key in choices:
                    if txt == value:
                        return key

                # Then try fuzzy match
                best_match, score = process.extractOne(txt, choices_txt, scorer=fuzz.ratio)
                if score > 80:
                    for value, key in choices:
                        if best_match == value:
                            return key

                return ''  # No match found

            except Exception as e:
                print(f"❌ Erro no mapeamento interno: {type(e).__name__} - {str(e)[:100]}")
                raise  # Re-raise to outer try-except

        try:
            # Normalize the text
            tipo_txt = ''.join(c for c in unicodedata.normalize('NFKD', tipo)
                               if not unicodedata.combining(c))

            categories = {
                'BEM': ['Material Esportivo', 'Uniformes'],
                'SERVICO': ['Recursos Humanos', 'Administrativa', 'Serviços', 'Identidades/Divulgações'],
                'OBRA': ['obra'],
                'TRIBUTO': ['tributo'],
                'OUTROS': ['']
            }

            tipo_gov = map_tipos_helper(tipo_txt, categories)

            if tipo_gov != '':
                return tipo_gov
            else:
                sys.exit()
        except Exception as e:
            print(f"❌ Erro no processamento do texto: {type(e).__name__} - {str(e)[:100]}")
            return None



def main() -> None:
    # Path to the .xlsx file that contains the necessary data to run the robot
    dir_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\SNEAELIS - Robô PAD')

    # Reference for the expense nature code
    cod_natureza_despesa = {
        'servicos': '33903999',
        'recursos_humanos': '33903999',
        'material': '33903014',
        'uniforme': '33903023',
        'impressos': '33903063',
        'premiacao': '33903004',
        'hidratacao_alimentacao': '33903007',
        'encargos_trab': '33903918',
        'material_esportivo': '33903014',
        'identidades/divulgações': '33903963',
        'identidades': '33903963',
        'divulgações': '33903963',
        'tributo': '33904718'
    }

    try:
        # Instantiate a Robo object
        robo = PWRobo(cod_natureza_despesa=cod_natureza_despesa)
        # Extract data from specific columns of the Excel file
    except Exception as e:
        print(f"\n‼️ Fatal error starting the robot: {e}")
        sys.exit("Stopping the program.")

    for root, dirs, files in os.walk(dir_path):
        for filename in files:
            if filename.endswith('.xlsx'):
                if root == dir_path:
                    path = os.path.join(root, filename)
                else:
                    path = root
                caminho_arquivo_fonte = os.path.join(root, filename)
                print(f"\n{'⚡' * 3}🚀 EXECUTING FILE: {filename} 🚀{'⚡' * 3}".center(70, '=')
                      , '\n')

                # DataFrame of the excel file
                df = robo.extrair_dados_excel(caminho_arquivo_fonte=caminho_arquivo_fonte)

                try:
                    pd.read_excel(caminho_arquivo_fonte, dtype=str, sheet_name='Status')
                    print(f'Sheet found!')
                except ValueError:
                    print(f'Sheet NOT found !')
                    print(f'Creating Sheet !')
                    # Create new status DataFrame if sheet doesn't exist
                    status_df = pd.DataFrame({
                        'Index': df.index,
                        'Status': [''] * len(df)
                    })
                    with pd.ExcelWriter(caminho_arquivo_fonte, engine='openpyxl', mode='a',
                                        if_sheet_exists='replace') as writer:
                        status_df.to_excel(writer, sheet_name='Status', index=False)

                # Start consultation and navigate to the process search page
                robo.consulta_proposta()

                numero_processo_temp = df.iloc[0,1]
                numero_processo = robo.fix_prop_num(numero_processo_temp)

                cnpj_xlsx = df.loc[df[0] == 'CNPJ:', 1].iloc[0]

                unique_values = []
                unique_values_col_b = df[1].unique()
                # first occurrence index
                unique_idx = np.where(unique_values_col_b == 'TIPO')[0][0]
                unique_values_temp = unique_values_col_b[unique_idx+1:]
                for val in unique_values_temp:
                    val = str(val).lower()
                    unique_values.append(val)

                status_df = pd.read_excel(caminho_arquivo_fonte, dtype=str, sheet_name='Status')
                for value in unique_values:
                    try:
                        if value == 'eventos' or value == 'alimentação':
                            continue
                        grouped_df = df[df[1].str.lower().str.contains(value, na=False)]

                        print(f"Executing for {value}\n"
                              f"Number of rows in grouped_df: {len(grouped_df)}"
                              .center(50, '-'), '\n')

                        if grouped_df.empty:
                            continue
                        for idx, row in grouped_df.iterrows():
                            # Read the guide spreadsheet, with marking of items already done
                            # Check if any line has already been executed
                            if status_df.iloc[idx, 1] == 'feito':
                                print(f'Line {idx} already executed. Skipping line\n')
                                continue
                            else:
                                break
                        else:
                            continue

                        robo.loop_de_pesquisa(df=grouped_df,
                                                  numero_processo=numero_processo,
                                                  tipo_desp=value,
                                                  cod_natur_desp=robo.map_cod_natur_desp(
                                                        dict_cod=cod_natureza_despesa,
                                                        cod=value
                                                        ),
                                                  cnpj_xlsx=cnpj_xlsx,
                                                  caminho_arquivo_fonte = caminho_arquivo_fonte
                                                  )
                    except BreakInnerLoop:
                        print("⚠️ Stopping this unique_values loop early.")
                        break
                    except KeyboardInterrupt:
                        print("Script stopped by user (Ctrl+C). Exiting cleanly.")
                        sys.exit(0) # Exit gracefully
                    except Exception as e:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(f"Error occurred at line: {exc_tb.tb_lineno}")
                        print(f"❌ Failed to execute script. Error: {type(e).__name__}\n{str(e)[:100]}")
                        sys.exit(0)  # Exit gracefully
                else:
                    print(path)
                    robo.logger.info(f'Successfully added PAD for proposal {numero_processo}, '
                                     f'deleting file {path}')
                    robo.delete_path(path)

if __name__ == "__main__":
    main()
