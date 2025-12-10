import os.path
import shutil
import time
import sys
import re
import unicodedata
import logging.handlers
import logging

import pandas as pd
import numpy as np

import autoit

import pytesseract

import pyautogui

from datetime import datetime

from colorama import  Fore

from PIL import Image
from io import BytesIO

from thefuzz import process, fuzz

from selenium import webdriver
from selenium.common import WebDriverException, TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains

class BreakInnerLoop(Exception):
    pass

class Robo:
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
            # Garante que nenhuma "Tab Search" seja aberta ao iniciar
            self.chrome_options.add_argument('--disable-features=TabSearch')
            self.chrome_options.add_argument('--disable-component-extensions-with-background-pages')

            # Configurações para download automático de PDF
            settings = {
                "plugins.always_open_pdf_externally": True,  # Abre PDF externamente
                "download.default_directory": self.caminho_download,
                "download.prompt_for_download": False,  # Não pergunta onde salvar
                "download.directory_upgrade": True,
                "safebrowsing.enabled": True,
                "profile.default_content_settings.popups": 0,  # Bloqueia pop-ups
                "profile.content_settings.exceptions.automatic_downloads.*.setting": 1,
            }
            self.chrome_options.add_experimental_option('prefs', settings)

            # Opções para melhor performance e evitar detecção
            self.chrome_options.add_argument('--start-maximized')
            self.chrome_options.add_argument('--disable-blink-features=AutomationControlled')
            self.chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
            self.chrome_options.add_experimental_option('useAutomationExtension', False)

            try:
                # Inicializa o driver do Chrome com as opções e o gerenciador de drivers
                self.driver = webdriver.Chrome(options=self.chrome_options)

            except Exception as e:
                print(f"Error with ChromeDriverManager: {e}")
                sys.exit()

            self.driver.switch_to.window(self.driver.window_handles[0])

            # Defines Logger
            self.logger = self.setup_logger()

            print(f"✅ Controle sobre a na página {self.driver.title}.")


        except WebDriverException as e:
            # Imprime mensagem de erro se a conexão falhar
            print(f"❌ Erro ao conectar ao navegador existente: {e}")
            # Define o driver como None em caso de falha na conexão
            self.driver = None

    def webdriver_element_wait(self, xpath: str, tm_ot: int=8):
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
            return WebDriverWait(self.driver, tm_ot).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        except Exception as e:
            raise e


    # Navega até a página de busca da proposta
    def consulta_proposta(self):
        """
               Navega pelas abas do sistema até a página de busca de processos.

               Esta função clica nas abas principal e secundária para acessar a página
               onde é possível realizar a busca de processos.
               """
        print(f'{'⚙️'*3}💼 INICIANDO CONSULTA DE PROPOSTAS 💼{'⚙️'*3}'.center(80, '='))
        print()
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
        print(f"{'🔎' * 3}🧭 ACESSANDO CAMPO DE PESQUISA 🧭{'🔎' * 3}".center(80, '='))
        print()

        try:
            # Seleciona campo de consulta/pesquisa, insere o número de proposta/instrumento e da ENTER
            self.driver.refresh()
            campo_pesquisa = self.webdriver_element_wait('//*[@id="consultarNumeroProposta"]')
            campo_pesquisa.clear()
            campo_pesquisa.send_keys(numero_processo)
            campo_pesquisa.send_keys(Keys.ENTER)
            try:
                # Acessa o item proposta/instrumento
                acessa_item = self.webdriver_element_wait('//*[@id="tbodyrow"]/tr/td[1]/div/a')
                acessa_item.click()
            except Exception as e:
                print(f' Processo número: {numero_processo}, não encontrado. Erro: {type(e).__name__}')
                self.logger.info(f'Processo número: {numero_processo}, não encontrado')
                raise BreakInnerLoop
        except Exception as e:
            print(f' Falha ao inserir número de processo no campo de pesquisa. Erro: {type(e).__name__}')


    def verif_reg(self, cnpj_xlsx: str):
        print(f"{'🗺️' * 3}📍 ACESSANDO EXTRATO DE REGULARIDADE 📍{'🗺️' * 3}".center(80, '='))
        print()
        try:
            # Navega para verificação de regularidade
            verif_area = self.webdriver_element_wait('//*[@id="menuInterno"]/div/div[11]')
            verif_area.click()
            # Clicka no link de acesso para a página de emissão de extratos
            verif_page = self.webdriver_element_wait('//*[@id="contentMenuInterno"]/div/div/ul/li[3]/a')
            verif_page.click()

            # Preenche cnjp e baixa o extrato
            cnjp_campo = self.webdriver_element_wait('//*[@id="consultarProponente:_idJsp22"]')
            cnjp_campo.clear()
            cnjp_campo.send_keys(cnpj_xlsx)
            cnjp_campo.send_keys(Keys.ENTER)

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Erro ao buscar endereço:{type(e).__name__}.\n Erro {str(e)[:100]}")


    # Insere o código na janela de seleção de município
    def cauc(self):
        print(f"{'🔎' * 3}🏙️ BUSCANDO DADOS CAUC 🏙️{'🔎' * 3}".center(80, '='))
        print()
        try:
            # Acessa aba requisitos
            aba_req = self.webdriver_element_wait('//*[@id="grupo_abas"]/a[3]')
            aba_req.click()

            # Acessa a área de pesquisa do CAUC
            aba_cauc = self.webdriver_element_wait('//*[@id="menu_link_2144784112_2061164"]')
            aba_cauc.click()

            # Faz a consulta mais recente do CAUC
            self.webdriver_element_wait('//*[@id="dados-cauc"]/siconv-historico/div/div[4]/button[1]').click()

            # Baixa documento PDF do CAUC
            tabela_ = self.webdriver_element_wait('//*[@id="listagem"]/tbody')
            consulta_ = tabela_.find_elements(By.TAG_NAME, './tr')[0]

            # Confere se a data da ultima pesquesa é a de hoje
            consulta_data = consulta_.find_elements(By.TAG_NAME, 'td')[0].text
            consulta_data = datetime.strptime(consulta_data.split()[0], '%d/%m/%Y').date()
            if consulta_data == datetime.now().date():
                celula_download = consulta_.find_elements(By.TAG_NAME, 'td')[4]
                celula_download.find_element(By.TAG_NAME, 'button').click()

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Falha ao cadastrar código do municipio"
                  f" {type(e).__name__}.\n Erro: {str(e)[:80]}")


    # Emite a certidão de regularidade do fgts
    def reg_fgts(self, cnpj: str, uf: str):
        # Abre nova aba e acessa o site para emiussão da certidão
        try:
            targe_url = r'https://consulta-crf.caixa.gov.br/consultacrf/pages/consultaEmpregador.jsf'
            self.driver.execute_script(f'window.open("{targe_url}", "_blank");')

            campo_cnpj = self.webdriver_element_wait('//*[@id="mainForm:txtInscricao1"]')
            campo_cnpj.clear()
            campo_cnpj.send_keys(cnpj)

            sel_uf =  self.webdriver_element_wait('//*[@id="mainForm:uf"]')
            sel_uf_obj = Select(sel_uf)
            sel_uf_obj.select_by_value(uf)

            # Clica em consultar
            self.webdriver_element_wait('//*[@id="mainForm:btnConsultar"]').click()

            # Acessa página do certificado
            self.webdriver_element_wait('//*[@id="mainForm:j_id51"]').click()

            # Abre o certificado
            self.webdriver_element_wait('//*[@id="mainForm:btnVisualizar"]').click()

            # Abre a tela de salvar em PDF
            self.webdriver_element_wait('//*[@id="mainForm:btImprimir4"]').click()

            # Salva o PDF na pasta Downloads
            autoit.win_wait("Imprimir", 3)
            autoit.control_click("Imprimir", "Button2")

            autoit.win_wait('Salvar como', 5)
            autoit.win_activate('Salvar como')
            time.sleep(0.3)
            autoit.control_click('Salvar como', 'Button1')

            self.driver.close()

            return True

        except Exception as e:
            print(f"Erro ao selecionar UF {uf}: {str(e)[:100]}")
            return False


    # Emite certião nada consta TST
    def certidao_tst(self, cnpj: str):
        # Abre nova aba e acessa o site para emissão da certidão TST
        try:
            targe_url = r'https://cndt-certidao.tst.jus.br/inicio.faces'
            self.driver.execute_script(f'window.open("{targe_url}", "_blank");')

            # Clica em emitir certidão
            self.webdriver_element_wait('//*[@id="corpo"]/div/div[2]/input[1]').click()

            # Preenche o campo de CNPJ ( usar pontos, barrars e traços)
            campo_cnpj = self.webdriver_element_wait('//*[@id="gerarCertidaoForm:cpfCnpj"]')
            campo_cnpj.clear()
            campo_cnpj.send_keys(cnpj)

            # Preenche captcha
            captcha = self.webdriver_element_wait('//*[@id="idImgBase64"]').screenshot_as_png
            img = Image.open(BytesIO(captcha))
            captcha_txt = pytesseract.image_to_string(img, config='--psm 8 --oem 3')
            print(captcha_txt)
            campo_captcha = self.webdriver_element_wait('//*[@id="idCampoResposta"]')
            campo_captcha.clear()
            campo_captcha.send_keys(captcha_txt)

            # Clica para emitir certidão
            self.webdriver_element_wait('//*[@id="gerarCertidaoForm:btnEmitirCertidao"]').click()

            self.driver.close()

            return True

        except Exception as e:
            print(f"Erro ao emitir certidão TST: {str(e)[:100]}")
            return False

    # Certidão CGU
    def certidao_cgu(self, cnpj) -> bool:
        # Abre nova aba e acessa o site para emiussão da certidão CGU
        try:
            targe_url = r'https://certidoes.cgu.gov.br/'
            self.driver.execute_script(f'window.open("{targe_url}", "_blank");')

            # Seleciona botão de rádio "Ente Privado"
            self.webdriver_element_wait('//*[@id="__BVID__82"]/div[1]').click()

            # Preenche o campo de CNPJ ( usar pontos, barrars e traços)
            campo_cnpj = self.webdriver_element_wait('//*[@id="cpfCnpj"]')
            campo_cnpj.clear()
            campo_cnpj.send_keys(cnpj)

            # Clica para fazer consulta
            self.webdriver_element_wait('//*[@id="consultar"]').click()

            # Clica para emitir certidão
            self.webdriver_element_wait('//*[@id="btnEmitirCertidao_8346688"]').click()

            self.driver.close()

            return True

        except Exception as e:
            print(f"Erro ao emitir certidão TST: {str(e)[:100]}")
            return False


    @staticmethod
    def extrair_dados_excel(caminho_arquivo_fonte):
        try:
            data_frame = pd.read_excel(caminho_arquivo_fonte,dtype=str,header=None,sheet_name=0)

            return data_frame
        except Exception as e:
            print(f"🤷‍♂️❌ Erro ao ler o arquivo excel: {os.path.basename(caminho_arquivo_fonte)}.\n"
                  f"Nome erro: {type(e).__name__}\nErro: {str(e)[:100]}")


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


    @staticmethod
    # Set's up logger
    def setup_logger(level=logging.INFO):
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


def main() -> None:
    # Caminho do arquivo .xlsx que contem os dados necessários para rodar o robô
    dir_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Robô PAD')
    try:
        # Instancia um objeto da classe Robo
        robo = Robo()
    except Exception as e:
        print(f"\n‼️ Erro fatal ao iniciar o robô: {e}")
        sys.exit("Parando o programa.")






if __name__ == "__main__":
    main()

