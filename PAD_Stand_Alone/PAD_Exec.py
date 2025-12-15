import os.path
import time
import sys
import re
import unicodedata
import shutil
import requests
import threading
import random
import subprocess

import pandas as pd
import numpy as np

import undetected_chromedriver as uc

import pyautogui

from pandas import ExcelWriter

from colorama import Fore

from thefuzz import process, fuzz

from selenium.common import (WebDriverException,
                             TimeoutException,
                             NoSuchElementException,
                             SessionNotCreatedException
                             )
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains



class BreakInnerLoop(Exception):
    pass

class Robo:
    # Inicia as principais instâncias do programa e define alguns objetos importantes
    def __init__(self, gui_callback=None):
        """
        Inicializa o objeto Robo, configurando e iniciando o driver do Chrome.
        """
        # Chrome setup
        self.chrome_process = None
        self.website_url = r'https://idp.transferegov.sistema.gov.br/idp/'

        # Captcha handling
        self.captcha_solved = False
        self.captcha_event = threading.Event()
        self.gui_callback = gui_callback

        try:
            # Inicializa o driver do Chrome
            self.driver = uc.Chrome(
                headless=False,
                version_main=None,
                options=self.create_chrome_options()
            )

            self.add_stealth_measures()

            print("✅ Navegador Chrome iniciado com sucesso.")
        except SessionNotCreatedException:
            chrome_version = self.get_chrome_version()
            driver_created = False

            if chrome_version:
                for version in [chrome_version, chrome_version-1, chrome_version+1]:
                    try:
                        self.driver = uc.Chrome(
                                headless=False,
                                version_main=version,
                                #options=self.create_chrome_options()
                            )
                        self.add_stealth_measures()
                        print(f"✅ Sucesso com versão {version}")
                        driver_created = True
                        break

                    except Exception as e:
                        print(f'Falha ao encontar uma versão que funcione.\nError type: {type(e).__name__} '
                              f'\nError: {str(e)[:200]}')
                        continue

                if not driver_created:
                    raise Exception("No compatible version found")

            else:
                raise Exception("Could not detect Chrome version")

        except WebDriverException as e:
            print(f"❌ Erro ao instanciar o driver.\nError type: {type(e).__name__}\nError: {str(e)[:200]}")
            sys.exit(1)
        except:
            print(f"❌ Erro fatal !")
            sys.exit(1)

    def add_stealth_measures(self):
        """Add additional stealth measures to avoid detection"""
        stealth_scripts = [
            # Remove webdriver property
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})",

            # Remove automation controlled property
            "Object.defineProperty(navigator, 'languages', {get: () => ['pt-BR', 'pt', 'en-US', 'en']})",

            # Mock permissions
            "const originalQuery = window.navigator.permissions.query; window.navigator.permissions.query = (parameters) => (parameters.name === 'notifications' ? Promise.resolve({ state: Notification.permission }) : originalQuery(parameters));",

            # Mock plugins
            "Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]})",

            # Mock hardware concurrency
            "Object.defineProperty(navigator, 'hardwareConcurrency', {get: () => 4})"
        ]

        for script in stealth_scripts:
            try:
                self.driver.execute_script(script)
            except:
                pass


    # Detecta a versão do Chrome
    def get_chrome_version(self):
        try:
            commands = [
                ['google-chrome', '--version'],
                ['chromium', '--version'],
                ['reg', 'query', 'HKEY_CURRENT_USER\\Software\\Google\\Chrome\\BLBeacon', '/v', 'version']
            ]

            for cmd in commands:
                try:
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
                    if result.returncode == 0:
                        version_match = re.search(r'(\d+)\.\d+\.\d+\.\d+', result.stdout)
                        if version_match:
                            return int(version_match.group(1))
                except:
                    continue

            return None

        except:
            return None


    def check_captcha(self):
        # Specific detection for the hCaptcha iframe you provided
        hcaptcha_selectors = [
            "//iframe[contains(@src, 'hcaptcha.com')]",
            "//iframe[contains(@title, 'hCaptcha')]",
            "//iframe[contains(@src, '93b08d40-d46c-400a-ba07-6f91cda815b9')]",  # Your specific sitekey
            "//div[@class='h-captcha']",
            "//div[contains(@id, 'hcaptcha')]",
            "//*[@data-sitekey='93b08d40-d46c-400a-ba07-6f91cda815b9']"  # Your sitekey
        ]

        for selector in hcaptcha_selectors:
            try:
                elements = self.driver.find_elements(By.XPATH, selector)
                if elements:
                    print(f"🔴 hCaptcha detected with selector: {selector}")
                    return True
            except:
                continue
        else:
            return False


    def handle_captcha(self):

        print("🔴 CAPTCHA detected - pausing execution...")

        # Notify GUI if callback exists
        if self.gui_callback:
            self.gui_callback('captcha_detected')

        # Wait for user to solve CAPTCHA (this pauses the script)
        print("⏳ Waiting for user to solve CAPTCHA...")
        self.captcha_event.wait()
        self.captcha_event.clear()

        print("🟢 CAPTCHA solved - resuming execution...")

        # Make sure the page is in good state
        WebDriverWait(self.driver, 7).until(
            lambda driver: driver.execute_script("return document.readyState") == "complete"
        )
        time.sleep(2)


    def resume_after_captcha(self):
        self.captcha_solved = True
        self.captcha_event.set()
        try:
            self.driver.switch_to.default_content()
        except:
            pass
        print("✅ Resume signal received from GUI")


    # Acessa o site das transferências discricionárias, na página de login
    def navegate_to_transfgov(self):
        """Navega para o website configurado após a inicialização do driver."""

        if not self.driver:
            print("❌ Driver não inicializado. Não é possível navegar para o website.")
            self.close()
            return False

        try:
            print(f"🌐 Navegando para: TransfereGov Discricionárias")
            self.driver.get(self.website_url)

            # Aguarda a página carregar
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/div[1]')))

            print("✅ Página carregada com sucesso.")
            return True

        except TimeoutException:
            print("❌ Timeout ao carregar a página. Verifique a conexão ou a URL.")
            return False
        except Exception as e:
            print(f"❌ Erro ao navegar para o website: {e}")
            return False


    # Fecha o navegador e limpa recursos.
    def close(self):
        if self.driver:
            try:
                self.driver.quit()
            except Exception as e:
                print(f"❌ Erro ao fechar navegador: {e}")
        if self.chrome_process:
            try:
                self.chrome_process.terminate()
            except:
                pass


    # Faz o login do site transferegov
    def log_in(self):
        def human_type(element, text):
            for char in text:
                element.send_keys(char)
                time.sleep(random.uniform(0.05, 0.22))

        def human_mouse_move(elmt):
            return False
            try:
                # Get element coordinates
                element_location = elmt.location
                element_size = elmt.size

                # Calculate the element's bounding box
                left
                y = element_location['y'] + element_size['height'] // 2

                # Add some randomness to avoid clicking exactly in center
                x += random.randint(-10, 10)
                y += random.randint(-10, 10)

                pyautogui.moveTo(x, y, duration=random.uniform(0.5, 1.2))

                time.sleep(random.uniform(0.05, 0.15))

                pyautogui.click()

                return True

            except:
                print(f"❌ PyAutoGUI click failed: {str(e)[:100]}")
                return False


        try:
            # Initiate login process
            log_in_btn = self.webdriver_element_wait('//*[@id="form_submit_login"]')
            if human_mouse_move(log_in_btn):
                print("✅ Human-like click performed")
            else:
                log_in_btn. click()

            # Fill in the credentials field
            credentials_field = self.webdriver_element_wait('//*[@id="accountId"]')
            human_type(element=credentials_field,text=self.__credentials)

            # Click on the continue button
            time.sleep(0.7)
            self.webdriver_element_wait('//*[@id="enter-account-id"]').click()

            if self.check_captcha():
                self.handle_captcha()

            self.driver.switch_to.default_content()

            psswd_field = self.webdriver_element_wait('//*[@id="password"]')
            human_type(element=psswd_field,text=self.__passcode)
            # Click to enter
            time.sleep(0.7)
            self.webdriver_element_wait('//*[@id="submit-button"]').click()

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(Fore.RED + f'🔴📄 Erro ao tentar fazer o login. \nErro: {e}')
            self.close()
            sys.exit(1)

    # Chama a função do webdriver com wait element to be clickable
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
            return WebDriverWait(self.driver, 8).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        except Exception as e:
            raise e


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
                #print(Fore.MAGENTA + "\n✅ Processo resetado com sucesso !")
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
            #print(f"{Fore.MAGENTA}✅ Sucesso em acessar a página de busca de processo{Style.RESET_ALL}")
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
            print(f' Falha ao inserir número de processo no campo de pesquisa. Erro: {type(e).__name__}')


    def  busca_endereco(self, cnpj_xlsx):
        cod_municipio_path = self.municipios_xslx_source('municipios.xlsx')

        # Debug the path
        print(f"🔍 Procurando arquivo municipios.xlsx em: {cod_municipio_path}")
        print(f"📁 Arquivo existe: {os.path.exists(cod_municipio_path)}")

        time.sleep(0.5)
        try:
            print(f'🔍 Iniciando busca de endereço'.center(50, '-'), '\n')

            # Aba Participantes
            self.webdriver_element_wait('/html/body/div[3]/div[14]/div[1]/div/div[2]/a[3]/div/span').click()
            # Botão detalhar
            time.sleep(0.5)
            self.webdriver_element_wait('//*[@id="form_submit"]').click()

            cnjp_web = self.webdriver_element_wait('//*[@id="txtCNPJ"]').text
            if cnpj_xlsx != cnjp_web:
                raise ValueError("CNPJ incompatível entre site e planilha")

            endereco = self.webdriver_element_wait('//*[@id="txtEndereco"]').text
            cep = endereco.split('CEP:')[-1].replace('-','')

            # acha o código do município
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
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Erro ao buscar endereço:{type(e).__name__}.\n Erro {str(e)[:100]}")


    # Insere o código na janela de seleção de município
    def cod_mun(self, cod_municipio):
        try:
            WebDriverWait(self.driver, 20).until(EC.number_of_windows_to_be(2))
            current_window = self.driver.current_window_handle
            windows = self.driver.window_handles
            for win in windows:
                if win != current_window:
                    self.driver.switch_to.window(win)
                    time.sleep(1)
            # Campo de inserção do número de código
            campo_prenchimento = self.webdriver_element_wait('//*[@id="consultarCodMunicipio"]')
            campo_prenchimento.clear()
            campo_prenchimento.send_keys(cod_municipio)
            campo_prenchimento.send_keys(Keys.ENTER)

            # Botão de consultar
            self.webdriver_element_wait('//*[@id="form_submit"]').click()
            self.webdriver_element_wait('/html/body/div[3]/div[3]/div[2]/table/tbody/tr/td[4]/nobr/a').click()

            # Volta para a DOM original
            self.driver.switch_to.window(current_window)
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Falha ao cadastrar código do municipio"
                  f" {type(e).__name__}.\n Erro: {str(e)[:80]}")


    def nav_plano_act_det(self):
        time.sleep(1)
        try:
            print(f'🚢 Navegando para o plano de ação detalhado:'.center( 50, '-'), '\n')
            # Volta para os dados da proposta
            self.webdriver_element_wait('//*[@id="lnkConsultaAnterior"]').click()

            # Aba Plano de trabalho
            self.webdriver_element_wait('//*[@id="div_997366806"]').click()

            # Seleciona Plano de Aplicação Detalhado
            self.webdriver_element_wait('//*[@id="menu_link_997366806_836661414"]/div').click()

        except Exception as e:
            print(f"❌ Ocorreu um erro ao executar ao pesquisar Plano de Ação Detalhado: {type(e).__name__}"
                  f".\n Erro {str(e)[:50]}")


    # Loop para adicionar PAD da proposta
    def loop_de_pesquisa(self, df, numero_processo: str, tipo_desp: str,
                         cod_natur_desp: str, cnpj_xlsx: str, caminho_arquivo_fonte: str):
        def sanitize_txt(txt:str) -> str:
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
            """
                Remove all integers from text (both leading and trailing)
                """
            return re.sub(r'^\d+\s*|\s*\d+$', '', txt).strip()


        # Inicia o processo de consulta do instrumento
        try:
            status_df = pd.read_excel(caminho_arquivo_fonte, dtype=str, sheet_name='Status')

            # Pesquisa pelo processo
            self.campo_pesquisa(numero_processo=numero_processo)

            # Faz a busca dos dados de localização
            endereco, cep, cod_municipio = self.busca_endereco(cnpj_xlsx=cnpj_xlsx)

            # Executa pesquisa de anexos
            self.nav_plano_act_det()

            # Seleciona lista de anexos execução e manda baixar os arquivos
            print('🌐 Acessando página de preenchimento PAD'.center( 50, '-'), '\n')
            time.sleep(1.5)
            tipo_serv = self.map_tipos(tipo_desp)

            # Locate the dropdown element
            dropdown = Select(self.driver.find_element("id", "incluirBemTipoDespesa"))

            # Select by visible text
            dropdown.select_by_value(f"{tipo_serv}")

            try:
                # Clica para incluir
                btn_incluir = self.driver.find_element(By.XPATH, '//*[@id="form_submit"]')
                if btn_incluir.get_attribute("value") == 'Incluir':
                    btn_incluir.click()
                else:
                    log_fechados = (
                        r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                        r'Social\SNEAELIS - Robô PAD\lista_pad_fechados.json')
                    self.salva_progresso(arquivo_log=log_fechados, nome_arquivo=numero_processo)
                    return False
            except Exception:
                print(f'Botão de incluir não localizado, pulando processo: {numero_processo}')
                return

            lista_campos = [
                # [0] Descrição do item
                '//*[@id="incluirBensDescricaoItem"]',
                # [1] Código da Natureza de Despesa
                '//*[@id="incluirBensCodigoNaturezaDespesa"]',
                # [2] Unidade Fornecimento
                '//*[@id="incluirBensCodUnidadeFornecimento"]',
                # [3] Valor Total
                '//*[@id="incluirBensValor"]',
                # [4] Quantidade
                '//*[@id="incluirBensQuantidade"]',
                # [5] Endereço de Localização
                '//*[@id="incluirBensEndereco"]',
                # [6] CEP
                '//*[@id="incluirBensCEP"]',
            ]

            print('📝 Preenchendo PAD'.center(50, '-'), '\n')

            for idx, row in df.iterrows():
                # Lê a planilha guia, com marcação dos itens já feitos
                # Verifica se alguma linha já foi executada
                if status_df.iloc[idx, 1] == 'feito':
                    print(f'Linha {idx} já executada. Pulando linha\n')
                    continue
                print(f"Preenchendo item nº:{idx} {str(row.iloc[2])}\n".center(50))
                try:
                    # Descrição do item
                    desc_item_txt = str(row.iloc[2]) + ': ' + str(row[3])
                    desc_item_txt = sanitize_txt(desc_item_txt)
                    desc_item = self.webdriver_element_wait(lista_campos[0])
                    desc_item.clear()
                    desc_item.send_keys(desc_item_txt)

                    # Código da Natureza de Despesa
                    cod_natur = self.webdriver_element_wait(lista_campos[1])
                    cod_natur.clear()
                    cod_natur.send_keys(cod_natur_desp)

                    # Unidade Fornecimento
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
                    elif un_fornecimento in ['litro']:
                        un_fornecimento = 'L'
                    elif un_fornecimento in ['pacote']:
                        un_fornecimento = 'PCT'
                    elif pd.isna(un_fornecimento) or un_fornecimento == 'nan' or un_fornecimento == 'N/A':
                        print('Valor não encontrado ou igual a zero\n')
                        continue

                    un_forn_field = self.webdriver_element_wait(lista_campos[2])
                    un_forn_field.clear()
                    un_forn_field.send_keys(un_fornecimento)

                    # Valor Total
                    valor_total = str(row.iloc[24])
                    if pd.isna(valor_total) or valor_total == 'nan' or valor_total == 'N/A':
                        print('Valor não encontrado ou igual a zero\n')
                        continue

                    if '.' in valor_total:
                        if len(valor_total.split('.')[-1]) == 1:
                            valor_total = valor_total + "0"
                    else:
                        valor_total = valor_total + "00"

                    valor_total_field = self.webdriver_element_wait(lista_campos[3])
                    valor_total_field.clear()
                    valor_total_field.send_keys(valor_total)

                    # Quantidade
                    qtd = str(row.iloc[9])
                    if pd.isna(valor_total) or valor_total == 'nan' or valor_total == 'N/A':
                        print('Valor não encontrado ou igual a zero\n')
                        continue
                    qtd = qtd + "00"
                    qtd_field = self.webdriver_element_wait(lista_campos[4])
                    qtd_field.clear()
                    qtd_field.send_keys(qtd)

                    # Endereço de Localização
                    end_loc = self.webdriver_element_wait(lista_campos[5])
                    end_loc.clear()
                    end_loc.send_keys(endereco)

                    # CEP
                    cep_element = self.webdriver_element_wait(lista_campos[6])
                    cep_element.clear()
                    cep_element.send_keys(cep.strip())

                    # Código do Município
                    self.driver.find_elements(By.CLASS_NAME, "btnBusca")[2].click()
                    self.cod_mun(cod_municipio)

                    # Botão "Incluir"
                    self.driver.find_elements(By.CSS_SELECTOR, "input#form_submit")[0].click()
                    time.sleep(2)

                    try:
                        # Try to find the CAPTCHA error
                        error_captcha = (WebDriverWait(self.driver, 1.2)
                                         .until(EC.visibility_of_element_located((By.CLASS_NAME, 'errors'))))
                        # If we get here, CAPTCHA error WAS found
                        print("CAPTCHA error detected - not marking as done")
                        if error_captcha:
                            sys.exit("Erro no preenchimento do PAD.\nTerminando execução")
                    except TimeoutException:
                        try:
                            if mark_as_done(status_df, idx, caminho_arquivo_fonte):
                                print("✅ Successfully marked and saved\n")
                            else:
                                print("❌ Failed to save\n")
                        except Exception as e:
                            print(f"❌ Error saving to Excel: {type(e).__name__}\n Erro {str(e)[:100]}\n")
                    except Exception as er:
                        print(f"❌ Falha {type(er).__name__}.\nErro{er}\n")
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(f"Error occurred at line: {exc_tb.tb_lineno}")
                    print(f"❌ Falha ao cadastrar PAD {type(e).__name__}.\n Erro: {str(e)[:80]}\n")
                    return False

            # Botão "Encerrar"
            time.sleep(1)
            self.driver.find_elements(By.CSS_SELECTOR, "input#form_submit")[1].click()
            time.sleep(1)
            self.consulta_proposta()
            return True

        except TimeoutException as t:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f'TIMEOUT {str(t)[:50]}')
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            self.consulta_proposta()
        except Exception as erro:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"❌ Error occurred at line: {exc_tb.tb_lineno}")
            print(f'❌ Falha ao tentar incluir documentos. Erro:{type(erro).__name__}\n'
                  f'{str(erro)[:100]}...')
            self.consulta_proposta()


    def map_cod_natur_desp(self,dict_cod: dict, cod: str, threshhold: int=80) -> str:
        norm = self.normaliza_text(cod)
        choices = dict_cod.keys()

        try:
            # Check if the string, after normilized, is compatible with any key
            if norm in dict_cod:
                return dict_cod.get(norm)

            # Check if the string has any similarity with some key, threshhold is 80%
            best, score = process.extractOne(query=norm, choices=choices, scorer=fuzz.ratio)
            if score < threshhold:
                print(f"\n‼️ Falha ao identificar código de natureza de despesa.")
                sys.exit("Parando o programa.")

            return dict_cod.get(best)

        except Exception as e:
            print(f"\n‼️ Erro fatal ao inserir PAD: {type(e).__name__}\nErro == {str(e)[:100]}")
            sys.exit("Parando o programa.")


    def main(self, files: list) -> None:
        # Referência para o código de natureza da despesa
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

        for idx, file in enumerate(files, start=1):
            print(f"▶️ ({idx}/{len(files)}) Processando: {file}")
            # DataFrame do arquivo excel
            df = self.extrair_dados_excel(caminho_arquivo_fonte=file)

            try:
                pd.read_excel(file, dtype=str, sheet_name='Status')
                print(f'Sheet found!')
            except ValueError:
                print(f'Sheet NOT found !')
                print(f'Creating Sheet !')
                # Create new status DataFrame if sheet doesn't exist
                status_df = pd.DataFrame({
                    'Index': df.index,
                    'Status': [''] * len(df)
                })
                with pd.ExcelWriter(file, engine='openpyxl', mode='a',
                                    if_sheet_exists='replace') as writer:
                    status_df.to_excel(writer, sheet_name='Status', index=False)

            # Navegate to desired website
            self.navegate_to_transfgov()

            # Initiate login
            self.log_in()

            # inicia consulta e leva até a página de busca do processo
            self.consulta_proposta()

            numero_processo_temp = df.iloc[0,1]
            numero_processo = self.fix_prop_num(numero_processo_temp)

            cnpj_xlsx = df.loc[df[0] == 'CNPJ:', 1].iloc[0]

            unique_values = []
            unique_values_col_b = df[1].unique()
            # first occurrence index
            unique_idx = np.where(unique_values_col_b == 'TIPO')[0][0]
            unique_values_temp = unique_values_col_b[unique_idx+1:]
            for val in unique_values_temp:
                val = str(val).lower()
                unique_values.append(val)

            status_df = pd.read_excel(file, dtype=str, sheet_name='Status')
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
                        # Lê a planilha guia, com marcação dos itens já feitos
                        # Verifica se alguma linha já foi executada
                        if status_df.iloc[idx, 1] == 'feito':
                            print(f'Linha {idx} já executada. Pulando linha\n')
                            continue
                        else:
                            break
                    else:
                        continue

                    self.loop_de_pesquisa(df=grouped_df,
                                              numero_processo=numero_processo,
                                              tipo_desp=value,
                                              cod_natur_desp=self.map_cod_natur_desp(
                                                    dict_cod=cod_natureza_despesa,
                                                    cod=value
                                                    ),
                                              cnpj_xlsx=cnpj_xlsx,
                                              caminho_arquivo_fonte = file
                                              )
                except BreakInnerLoop:
                    print("⚠️ Stopping this unique_values loop early.")
                    break
                except KeyboardInterrupt:
                    print("Script stopped by user (Ctrl+C). Exiting cleanly.")
                    self.close()
                    sys.exit(0) # Exit gracefully
                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(f"Error occurred at line: {exc_tb.tb_lineno}")
                    print(f"❌ Falha ao executar script. Erro: {type(e).__name__}\n{str(e)[:100]}")
                    self.close()
                    sys.exit(0)  # Exit gracefully
            else:
                self.delete_path(file)
        else:
            self.close()
            sys.exit(0)  # Exit gracefully

    @staticmethod
    def municipios_xslx_source(relative_path):
        """Get absolute path to resource, works for dev and for PyInstaller"""
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS

        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)


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
    def extrair_dados_excel(caminho_arquivo_fonte):
        try:
            data_frame = pd.read_excel(caminho_arquivo_fonte,dtype=str,header=None,sheet_name=0)

            return data_frame
        except Exception as e:
            print(f"🤷‍♂️❌ Erro ao ler o arquivo excel: {os.path.basename(caminho_arquivo_fonte)}.\n"
                  f"Nome erro: {type(e).__name__}\nErro: {str(e)[:100]}")


    @staticmethod
    # Mapeia os tipos de despesas nos quatro tipos disponíveis no transferegov
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


    @staticmethod
    def session_status(driver_service_url='http://localhost:9222'):
        try:
            response = requests.get(f"{driver_service_url}/sessions")
            if response.status_code == 200:
                sessions = response.json().get('value', [])
                print(f"Active browser sessions on port: {len(sessions)}")
                for session in sessions:
                    print(f"Session ID: {session.get('id')}, Capabilities: {session.get('capabilities')}")
            else:
                print(f"Unable to fetch sessions: Status {response.status_code}")
        except Exception as e:
            print(f"Error checking sessions: {str(e)[:100]}")


    # Create a new instance of Chrome Options
    @staticmethod
    def create_chrome_options():
        chrome_options = uc.ChromeOptions()

        # Suas configurações existentes
        chrome_options.add_argument("--no-first-run")
        #chrome_options.add_argument("--incognito")
        chrome_options.add_argument("--no-default-browser-check")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-blink-features")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-plugins")
        chrome_options.add_argument("--start-maximized")

        return chrome_options

    # Delete file from specified directory
    @staticmethod
    def delete_path(path:str):
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
# 888