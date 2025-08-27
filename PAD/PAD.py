import os.path
import time
import sys
import json
import re
import unicodedata

import pandas as pd
import numpy as np

import requests

from typing import List

from colorama import  Fore, Style
from rapidfuzz.distance.DamerauLevenshtein_py import similarity

from thefuzz import process, fuzz

from selenium import webdriver
from selenium.common import WebDriverException, TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver import ActionChains


class Robo:
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
        cod_municipio_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e '
                              r'Assistência Social\Teste001\municipios.xlsx')
        time.sleep(1)
        try:
            print(f'🔍 Iniciando busca de endereço'.center(50, '-'), '\n')

            # Aba Participantes
            self.webdriver_element_wait('/html/body/div[3]/div[14]/div[1]/div/div[2]/a[3]/div/span').click()
            # Botão detalhar
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

    # Mapeia os tipos de despesas nos quatro tipos disponíveis no transferegov
    def map_tipos(self, tipo):
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


    def session_status(self,driver_service_url='http://localhost:9222'):
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


    # Loop para adicionar PAD da proposta
    def loop_de_pesquisa(self, df, numero_processo: str, arquivo_log: str, tipo_desp: str,
                         cod_natur_desp: str, cnpj_xlsx: str):
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
        # Inicia o processo de consulta do instrumento
        try:
            # Pesquisa pelo processo
            self.campo_pesquisa(numero_processo=numero_processo)

            # Faz a busca dos dados de localização
            endereco, cep, cod_municipio = self.busca_endereco(cnpj_xlsx=cnpj_xlsx)

            # Executa pesquisa de anexos
            self.nav_plano_act_det()

            # Seleciona lista de anexos execução e manda baixar os arquivos
            print('🌐 Acessando página de preenchimento PAD'.center( 50, '-'), '\n')
            time.sleep(1)
            tipo_serv = self.map_tipos(tipo_desp)

            # Locate the dropdown element
            dropdown = Select(self.driver.find_element("id", "incluirBemTipoDespesa"))

            # Select by visible text
            dropdown.select_by_value(f"{tipo_serv}")

            # Clica para incluir
            self.driver.find_element(By.XPATH, '//*[@id="form_submit"]').click()

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
                print(f"Preenchendo item nº:{idx} {str(row.iloc[2])}\n".center(50))
                try:
                    # Descrição do item
                    desc_item_txt = str(row.iloc[2]) + '\n' + str(row[3])
                    desc_item_txt = sanitize_txt(desc_item_txt)
                    self.webdriver_element_wait(lista_campos[0]).send_keys(desc_item_txt)

                    # Código da Natureza de Despesa
                    self.webdriver_element_wait(lista_campos[1]).send_keys(cod_natur_desp)

                    # Unidade Fornecimento
                    un_fornecimento = str(row.iloc[8]).strip().lower()
                    if un_fornecimento == 'mensal':
                        un_fornecimento = 'MÊS'
                    elif un_fornecimento == 'unidade':
                        un_fornecimento = 'UN'
                    elif un_fornecimento == 'diaria' or un_fornecimento == 'diária':
                        un_fornecimento = 'DIA'
                    self.webdriver_element_wait(lista_campos[2]).send_keys(un_fornecimento)

                    # Valor Total
                    valor_total = str(row.iloc[23])
                    if '.' in valor_total:
                        self.webdriver_element_wait(lista_campos[3]).send_keys(valor_total)
                    else:
                        valor_total = str(row.iloc[23]) + "00"
                        self.webdriver_element_wait(lista_campos[3]).send_keys(valor_total)

                    # Quantidade
                    qtd = str(row.iloc[9])+"00"
                    self.webdriver_element_wait(lista_campos[4]).send_keys(qtd)

                    # Endereço de Localização
                    self.webdriver_element_wait(lista_campos[5]).send_keys(endereco)

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

                except Exception as e:
                    exc_type, exc_value, exc_tb = sys.exc_info()
                    print(f"Error occurred at line: {exc_tb.tb_lineno}")
                    print(f"❌ Falha ao cadastrar PAD {type(e).__name__}.\n Erro: {str(e)[:80]}")
                    sys.exit()
            # Botão "Encerrar"
            self.driver.find_elements(By.CSS_SELECTOR, "input#form_submit")[1].click()
            time.sleep(1)
            self.consulta_proposta()
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
                #print(f'Botão localizado com o seletor {locator[0]}')
                return self.driver.find_element(*locator)
            except Exception as e:
                print(f"Falha com localizador {locator}: {str(e)[:80]}")
                continue

        raise NoSuchElementException("Could not find button using any locator strategy")


    # Acessa a lista de anexos execução
    def lista_execucao(self) -> None:
        """
               Acessa a lista de anexos da execução.

               Esta função clica no botão para exibir a lista de anexos da execução
               e define o nome da coluna que contém as datas dos anexos.

               Returns:
                   None
               """
        try:
            # Seleciona lista de anexos execução e acessa a mesma
            lista_anexos_execucao = self.webdriver_element_wait('//tbody//tr//input[2]')
            lista_anexos_execucao.click()
            # Define o nome da coluna de data
        except TimeoutException:
            print(f"Timeout waiting for element: {'//tbody//tr//input[2]'}")
        except Exception as e:  # Catch other potential exceptions
            print(f"🤷‍♂️❌ Erro ao tentar entra na lista de anexos execução: {e}")


    def extrair_dados_excel(self, caminho_arquivo_fonte):
        try:
            data_frame = pd.read_excel(caminho_arquivo_fonte, dtype=str, header=None)

            return data_frame
        except Exception as e:
            print(f"🤷‍♂️❌ Erro ao ler o arquivo excel: {os.path.basename(caminho_arquivo_fonte)}.\n"
                  f"Nome erro: {type(e).__name__}\nErro: {str(e)[:100]}")


    # Salva o progresso em um arquivo json
    def salva_progresso(self, arquivo_log: str, nome_arquivo: str):
        """
        Salva o progresso atual em um arquivo JSON.
        :param arquivo_log: Endereço do arquivo JSON.
        :param nome_arquivo: O nome do arquivo que foi executado.
        """

        # Carrega os dados antigos
        lista_arquivos = self.carrega_progresso(arquivo_log=arquivo_log)

        # Verifica se o arquivo já está na lista para evitar duplicatas
        if nome_arquivo not in lista_arquivos:
            lista_arquivos.append(nome_arquivo)

        with open(arquivo_log, 'w', encoding='utf-8') as arq:
            json.dump(lista_arquivos, arq, indent=4, ensure_ascii=False)
        print("💾 Progresso salvo")


    # Carrega os dados do arquivo JSON que sereve como Cartão de Memória
    def carrega_progresso(self, arquivo_log: str) -> List[str]:
        """
            Carrega o progresso do arquivo JSON.
            :param arquivo_log: Endereço do arquivo JSON
            :return: Uma lista de nomes de arquivos. Se o arquivo não existir,
             estiver vazio ou inválido, retorna uma lista vazia.
        """
        if not os.path.exists(arquivo_log):
            return []

        try:
            with open(arquivo_log, 'r', encoding='utf-8') as arq:
                dados = json.load(arq)
                # Garante que o conteúdo seja uma lista
                if isinstance(dados, list):
                    return dados
                else:
                    return []
        except (json.JSONDecodeError, FileNotFoundError):
            # Retorna uma lista vazia em caso de erro na leitura do arquivo
            return []


    # Corrige o número da proposta que vem na planilha
    def fix_prop_num(self,numero_proposta):
        if '_' in numero_proposta:
            numero_proposta_fixed = numero_proposta.replace('_', '/')
        else:
            numero_proposta_fixed = numero_proposta
        return numero_proposta_fixed


    def normaliza_text(self, txt: str) -> str:
        if txt is None:
            return ''

        normal_text = ''.join(char for char in unicodedata.normalize('NFKD', txt) if not
        unicodedata.combining(char))

        normal_text = re.sub(r'[^a-z0-9]+', '_', normal_text)
        normal_text = re.sub(r'_+', '_', normal_text).strip()

        return normal_text


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

    def create_pad_file(self, dir_path: str) -> bool:
        """
        Creates a file named 'PAD Feito.txt' in the specified directory.

        Args:
            dir_path (str): The path to the directory where the file will be created.

        Returns:
            bool: True if file creation was successful, False otherwise.
        """
        try:
            # Ensure the directory path ends with a separator
            directory_path = os.path.normpath(dir_path, )

            # Create full file path
            file_path = os.path.join(directory_path, "PAD Feito.txt")

            # Check if directory exists, create it if it doesn't
            if not os.path.exists(directory_path):
                os.makedirs(directory_path)

            # Create or overwrite the file
            with open(file_path, 'w') as f:
                f.write("")
            return True

        except Exception as e:
            print(f"Error creating file: {str(e)[:100]}")
            return False

def main() -> None:
    # Caminho do arquivo JSON que serve como catão de memória
    arquivo_log = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                   r'Social\Teste001\PAD\lista_exec_pad.json')
    # Caminho do arquivo .xlsx que contem os dados necessários para rodar o robô
    dir_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                r'Social\SNEAELIS - Robô PAD')
    try:
        # Instancia um objeto da classe Robo
        robo = Robo()
        # Extrai dados de colunas específicas do Excel
    except Exception as e:
        print(f"\n‼️ Erro fatal ao iniciar o robô: {e}")
        sys.exit("Parando o programa.")

    for root, dirs, files in os.walk(dir_path):
        for filename in files:
            if filename.endswith('.xlsx'):
                lista_arq_exec = robo.carrega_progresso(arquivo_log)
                if filename in lista_arq_exec:
                    continue
                caminho_arquivo_fonte = os.path.join(root, filename)
                print(f"\nExecuting file: {filename}".center(50, '-'), '\n')

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
                    'divulgações': '33903963'
                }

                # DataFrame do arquivo excel
                df = robo.extrair_dados_excel(caminho_arquivo_fonte=caminho_arquivo_fonte)

                # inicia consulta e leva até a página de busca do processo
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
                        robo.loop_de_pesquisa(df=grouped_df,
                                              arquivo_log=arquivo_log,
                                              numero_processo=numero_processo,
                                              tipo_desp=value,
                                              cod_natur_desp=robo.map_cod_natur_desp(
                                                    dict_cod=cod_natureza_despesa,
                                                    cod=value
                                                    ),
                                              cnpj_xlsx=cnpj_xlsx
                                              )
                    except KeyboardInterrupt:
                        print("Script stopped by user (Ctrl+C). Exiting cleanly.")
                        sys.exit(0) # Exit gracefully
                    except Exception as e:
                        exc_type, exc_value, exc_tb = sys.exc_info()
                        print(f"Error occurred at line: {exc_tb.tb_lineno}")
                        print(f"❌ Falha ao executar script. Erro: {type(e).__name__}\n{str(e)[:100]}")
                        sys.exit(0)  # Exit gracefully

                robo.salva_progresso(arquivo_log=arquivo_log, nome_arquivo=filename)
                robo.create_pad_file(os.path.join(root))

if __name__ == "__main__":
    start_time = time.time()

    main()

    end_time = time.time()
    tempo_total = end_time - start_time
    horas = int(tempo_total // 3600)
    minutos = int((tempo_total % 3600) // 60)
    segundos = int(tempo_total % 60)
    print(f'⏳ Tempo de execução: {horas}h {minutos}m {segundos}s')