#from pandas.core.config_init import pc_max_info_cols_doc
from selenium import webdriver
from selenium.common import WebDriverException, TimeoutException, NoSuchElementException, \
    InvalidSelectorException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
#from selenium.webdriver.support.expected_conditions import number_of_windows_to_be
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
from tqdm import tqdm
from colorama import init, Fore, Back, Style
from functools import wraps
import pandas as pd
import time
import os
import sys
import shutil
import json


class Robo:
    def __init__(self):
        """
        Inicializa o objeto Robo, configurando e iniciando o driver do Chrome.
        """
        try:
            # Configuração do registro
            #self.arquivo_registro = ''
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

    # Pega o ultimo arquivo baixado da pasta Downloads e move para a pasta destino
    def espera_completar_download (self, pasta_download: str, tempo_limite=30):
        """
        Aguarda a conclusão de um download verificando a ausência de arquivos temporários ".crdownload".

        Parâmetros:
        -----------
        pasta_download : str
            Caminho da pasta onde os arquivos estão sendo baixados.

        tempo_limite : int, opcional (padrão: 30)
            Tempo máximo (em segundos) que a função aguardará antes de levantar uma exceção.

        nome_arquivo : str, opcional (padrão: 'None')
            Nome do arquivo esperado para download. Apenas utilizado na mensagem de erro.

        Retorna:
        --------
        bool
            Retorna `True` assim que todos os downloads forem concluídos (ou seja, não houver arquivos `.crdownload` na pasta).

        Levanta:
        --------
        Exception
            Se o tempo limite for atingido e ainda houver arquivos `.crdownload`, uma exceção é lançada com uma mensagem de erro.

        Descrição:
        ----------
        - A função inicia um temporizador que expira após `tempo_limite` segundos.
        - Em um loop contínuo, verifica se existem arquivos com a extensão `.crdownload` na pasta de downloads.
        - Se não houver arquivos `.crdownload`, retorna `True`, indicando que o download foi concluído.
        - Caso contrário, aguarda 1 segundo antes de verificar novamente.
        - Se o tempo limite for atingido e ainda existirem `.crdownload`, uma exceção é lançada.
        """
        # Inicia comparador de tempo
        tempo_final = time.time() + tempo_limite
        while time.time() < tempo_final:
            # Checa se existem arquivos com o placeholder .crdownload
            arq_crdownload = [f for f in os.listdir(pasta_download) if f.endswith('.crdownload') or
                              f.endswith('.tmp')]
            if not arq_crdownload:
                return True
        raise TimeoutException(f'Não completou o download no tempo limite')

    # Pega a data do documento e o nome da aba
    def compara_data(self, data_site: str) -> bool:
        """
        Compara a data fornecida com a data de ontem.

        Args:
            data_site: Data a ser comparada, no formato 'AAAA-MM-DD' ou 'AAAA-MM-DD HH:MM:SS'.

        Returns:
            True se a data fornecida for anterior à data de ontem, False caso contrário.
            Retorna False se o formato da data fornecida for inválido.
        """

        # Obtém a data de ontem
        try:
            # Converte a data do site para datetime, tratando diferentes formatos
            data_site_dt = self.converter_data(data_site)
            if data_site_dt is None:
                return False  # Formato de data inválido
            if datetime.isoweekday(datetime.today()) == 1:
                data_ontem = datetime.now() - timedelta(days=4)
            else:
                data_ontem = datetime.now() - timedelta(days=2)
            # Comparação
            if data_site_dt >= data_ontem:
                return True
            else:
                return False

        except Exception:
            print(f"❌ Erro na comparação de datas")
            return False

    # Converte o formato da data para que possa ser comparada
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

        print(f"Formato de data inválido: {data}")
        return None

    # pega a data do dia que está executando o programa
    def data_hoje(self):
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
            try:
                if '.' in nome_arq and not nome_arq.startswith('.'):
                    base, ext = nome_arq.rsplit('.', 1)
                else:
                    base = nome_arq
                    ext = ''
                # Verirfica o tamanho do nome do arquivo para que não dê erro de transferência
                if len(base) > 80:
                    novo_nome = f'Documento ({cont}).{ext}' if ext else f'Documento({cont})'
                    cont += 1
                else:
                    novo_nome = base.replace(' ', '').split()
                    novo_nome = f'{novo_nome}.{ext}' if ext else novo_nome
                if novo_nome != nome_arq:
                    new_filepath = os.path.join(pasta_download, novo_nome)
                    os.rename(old_filepath, new_filepath)
            except OSError as e:
                print(f"Error renaming '{nome_arq}': {e}")
                return None

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
            - Se o arquivo foi modificado **hoje ou depois**, move-o para `caminho_pasta`.
            - Exibe uma mensagem no console para cada arquivo movido.
            """
        # pega a data do dia que está executando o programa
        data_hoje = self.data_hoje()
        tempo_final = time.time() + tempo_limite
        self.permuta_nome_arq(pasta_download=pasta_download)
        while time.time() < tempo_final:
            try:
                for arq in os.listdir(pasta_download):
                    caminho_arq = os.path.join(pasta_download, arq)
                    # Data de modificação do arquivo
                    data_mod = datetime.fromtimestamp(os.path.getmtime(caminho_arq))
                    # Compara a data de hoje com a data de modificação do arquivo
                    if data_mod >= data_hoje:
                        # Move o arquivo para a pasta destino
                        shutil.move(caminho_arq, os.path.join(caminho_pasta, arq))
                print(f'📂🚚 Arquivos Movimentados !')
                if not os.listdir(pasta_download):
                    break
            except Exception as e:
                print(f'❌ Falha ao mover arquivo {e}')

    # Cria uma pasta com o nome especificado no one-drive e retorna o caminho.
    def criar_pasta(self, nome_pasta: str, caminho_pasta_onedrive: str) -> str:
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
        nome_pasta = nome_pasta
        caminho_pasta = os.path.join(caminho_pasta_onedrive, nome_pasta.replace('/', '_'))

        try:
            # Cria o diretório, incluindo o diretório pai, se necessário.
            exist_ok=True #evita erro se a pasta já existir.
            os.makedirs(caminho_pasta, exist_ok=True)
            print(f"📂✨ Pasta '{nome_pasta}' criada em: {caminho_pasta}")
        except Exception as e:
            print(f"❌ Erro ao criar a pasta '{nome_pasta}': {e}")
        # Retorna o caminho completo da pasta, mesmo que a criação tenha falhado (para tratamento posterior)
        return caminho_pasta

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
            i_limpo = i.replace('\n', '').replace('\r','').strip()
            # Anexa à lista limpa
            lista_limpa.append(i_limpo)
        return lista_limpa

    def extrair_dados_excel(self, caminho_arquivo_fonte, busca_id: str, status: str, links: str=None):
        """
           Lê os contatos de uma planilha Excel e executa ações baseadas nos dados extraídos.

           Args:
               caminho_arquivo_fonte (str): Caminho do arquivo Excel que será lido.
               busca_id (str): Nome da coluna que contém os números de processo.
               tipo_instrumento_id (str): Nome da coluna que contém o tipo de processo.
           """
        dados_processo = pd.read_excel(caminho_arquivo_fonte, dtype=str)
        #dados_processo.replace(u'\xa0', '', regex=True).infer_objects(copy=True)


        # Cria um lista para cada coluna do arquivo xlsx
        numero_processo = list()
        stt = list()
        links_list = list()
        try:
            # Itera a planilha e armazena os dados em listas
            for indice, linha in dados_processo.iterrows():  # Assume que a primeira linha e um cabeçalho
                numero_processo.append(linha[busca_id])  # Busca o número do processo
                stt.append(linha[status])
                if links:
                    links_list.append(linha[links])

        except Exception as e:
            print(f"❌ Erro de leitura encontrado, erro: {e}")
        if links:
            return numero_processo, stt, links_list

        return numero_processo, stt

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
        WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, xpath)))
        return WebDriverWait(self.driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath)))

    # Consulta o instrumento ou proposta (faz todo o caminho até chegar no item desejado)
    def consulta_instrumento(self):
        """
               Navega pelas abas do sistema até a página de busca de processos.

               Esta função clica nas abas principal e secundária para acessar a página
               onde é possível realizar a busca de processos.
               """
        # Reseta para página inicial
        try:
            reset = self.webdriver_element_wait('//*[@id="header"]')
            if reset:
                img = reset.find_element(By.XPATH,'//*[@id="logo"]/a/img')
                action = ActionChains(self.driver)
                action.move_to_element(img).perform()
                reset.find_element(By.TAG_NAME, 'a').click()
                print(Fore.MAGENTA + "\n✅ Processo resetado com sucesso !")
        except Exception as e:
            print(Fore.RED + f'🔄❌ Falha ao resetar.')

        # [0] Propostas; [1] Seleção PAC
        xpaths = ['//*[@id="menuPrincipal"]/div[1]/div[3]',
                  '//*[@id="contentMenu"]/div[2]/ul/li[4]/a'
                  ]
        try:
            for idx in range(len(xpaths)):
                self.webdriver_element_wait(xpaths[idx]).click()
            print(Fore.MAGENTA + "✅ Sucesso em acessar a página de busca de processo")
        except Exception as e:
            print(Fore.RED + f'🔴📄 Instrumento indisponível. \nErro: {e}')
            sys.exit(1)

    def pesquisa_processo(self, numero_processo:str):
        try:
            # Seleciona campo de consulta/pesquisa, insere o número de proposta/instrumento e da ENTER
            campo_pesquisa = self.webdriver_element_wait('//*[@id="formListarPropostaPac:_idJsp22"]')
            campo_pesquisa.clear()
            campo_pesquisa.send_keys(numero_processo)
            campo_pesquisa.send_keys(Keys.ENTER)
            print(f'\n{Fore.CYAN}🔎✔️ Campo pesquisado com sucesso{Style.RESET_ALL}')
            time.sleep(0.2)

        except Exception:
            print(f'{Fore.RED}❌🔍 Falha ao pesquisar processo: {numero_processo}{Style.RESET_ALL}')
            self.consulta_instrumento()

    # Recebe um XPATH e dá o comanod de click
    def clicar(self, xpath: str):
        try:
            click = self.webdriver_element_wait(xpath)
            if click:
                print(Fore.GREEN +f"⚠🖱️ Clicando em {click.text}")
                click.click()
                time.sleep(0.2)
            return False
        except Exception:
            print(
            Fore.RED + f"\⚠️🖱️ Erro ao tentar clickar no item: {self.webdriver_element_wait(xpath).text}")
            return True

    # Conta o número de páginas de tabelas o instrumento tem
    def conta_paginas(self, tabela) -> int:
        # Diz quantas páginas tem
        try:
            print(f'🔍📄 Buscando número de páginas')
            paginas = tabela.find_element(By.CSS_SELECTOR, '.pagelinks').text
            paginas = paginas.split(' ')[1]
            paginas = [int(num.strip()) for num in paginas.split(',')][-1]
            print(f'📄🔢Total de páginas encontradas:{paginas}\n')
            return paginas
        except NoSuchElementException:
            paginas = 1
            return paginas

    # Conta as páginas do objeto detalhado
    def conta_paginas_detalhe(self, tabela) -> int:
        # Diz quantas páginas tem
        try:
            print(f'🔍📄 Buscando número de páginas')
            paginas = tabela.find_element(By.CLASS_NAME, 'pagelinks').text
            paginas = paginas.split('(')[0]
            paginas.strip()
            paginas = int(paginas[-3: -1])
            print(f'📄🔢Total de páginas encontradas:{paginas}\n')
            return paginas
        except NoSuchElementException:
            paginas = 1
            return paginas


    # Passa para a próxima página
    def proxima_pag(self, pagina: int, max_pagina: int):
        try:
            if pagina <= max_pagina:
                print(Fore.GREEN + f'Passando para a página:{pagina}')
                time.sleep(0.3)
                self.driver.find_element(By.LINK_TEXT, f'{pagina}').click()
        except Exception as e:
            print(f'❌ Erro ao passar de página {e}')

    # Seleciona o webelement da tabela
    def acha_tabela(self, nome_tabela: str, mute=False, time=0.7):
        """Encontra um elemento de tabela usando múltiplas estratégias de localização.

            Args:
                nome_tabela: Pode ser um ID, XPath ou seletor CSS

            Returns:
                WebElement: O elemento da tabela encontrado

            Raises:
                TimeoutException: Se o elemento não for encontrado com nenhuma estratégia
                NoSuchElementException: Se o elemento não existir
                Exception: Para outros erros inesperados
            """
        locator_strategies = [
            (By.ID, "ID"),
            (By.XPATH, "XPath"),
            (By.CSS_SELECTOR, "CSS Selector")
                            ]
        try:
            for locator, strategy_name in locator_strategies:
                try:
                    if mute:
                        print(f"🔍 Buscando tabela pelo {strategy_name}")
                    return WebDriverWait(self.driver, time).until(
                        EC.element_to_be_clickable((locator, nome_tabela)))
                except TimeoutException:
                    print('Trying again')
                    continue
        except InvalidSelectorException:
            print(Fore.RED +f'📊❌ Tabela não encontrada usando nenhuma estratégia.'
                            f' Valor usado: {nome_tabela}')
        except Exception as e:
               print(Fore.RED +f"📊❌ Tabela não encontrada usando nenhuma estratégia. Valor usado: "
                               f"'{nome_tabela}'\nErro:{e}")

    # Acha as linhas da tabela
    def linhas(self, tabela):
        try:
            linhas = tabela.find_elements(By.TAG_NAME, 'tr')
            return linhas
        except Exception:
            print('❌ Falha ao tentar achar as linhas')

    # Entra no detalhamento do objeto
    def detalhar(self, linha, xpath, processo):
        try:
            # Pega o elemento do botão de "Detalhar" e deixa selecionado
            detalhar = linha.find_element(By.XPATH, xpath)
            texto = detalhar.text
            if detalhar:
                # Clica no botão "Detalhar"
                print(f'💻📄 Processo: {texto} acessado')
                detalhar.click()
        except Exception as e:
            print(f'{Fore.RED}❌📄 Falha ao acessar o processo: {processo}{Style.RESET_ALL}')
            print(f'Erro: {e}')

    # Baixa os arquivos
    def baixa_arquivo(self, linha, linha_num):
            try:
                # Pega o elemento do botão de baixar e deixa selecionado
                botao_download = linha.find_element(By.CSS_SELECTOR, 'a.buttonLink')
                # Pega o texto que nomeia o documento que será baixado, caso exita
                nome_arquivo = linha.find_element(By.CSS_SELECTOR,
                                                  "td[style='width:30%;']").text
                if botao_download:
                    # Clicka no botão "Baixar"
                    botao_download.click()
                    return  nome_arquivo
            except NoSuchElementException:
                print(f'{Fore.RED}📥❌ Botão de download não encontrado.'
                      f'\nNúmero da linha: {linha_num}{Style.RESET_ALL}')
            except Exception as e:
                print(f'Erro: {e}')

    # Confirma se existe um elemento
    def confirma(self, xpath: str):
        try:
           conf = WebDriverWait(self.driver, 3).until(EC.
                                                       presence_of_element_located((By.XPATH,xpath)))
           if conf:
               return True
        except Exception:
            print("\033[31m🔍🐛🚫 Impossível localizar o elemento.")
            return False

    # Salva o progresso em um arquivo json
    def salva_progresso(self, arquivo_log: str, processo_visitado: str, indice: int):
        """
        Salva o progresso atual em um arquivo JSON.

        :param arquivo_log: Endereço do arquivo JSON
        :param processo_visitado: Lista de processos concluídos.
        :param indice: Diz qual linha o programa iterou por último.
        """
        # Carrega os dados antigos
        dados_log = self.carrega_progresso(arquivo_log)

        # Carrega os dados novos
        novo_item = {
            "processo_visitado": processo_visitado,
            "indice": indice
        }

        dados_log.update(novo_item)
        with open(arquivo_log, 'w', encoding='utf-8') as arq:
            json.dump(dados_log, arq, indent=4)
        print(f"\n💾 Progresso salvo no processo:{processo_visitado} e ìndice: {dados_log["indice"]}")

    # Salva os processos com erro.
    def salva_erro(self, numero_processo):
        with open(r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                  r'Social\Power BI\Python\Requerimento de serviço\CGAG(Cássia)\log_erros.txt', 'a'
          ) as file:
            file.write(f'{numero_processo}\n')
        print(f'💾📑❌ Processo: {numero_processo} com erro, salvo')

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
            "processo_visitado": 000000,
            "indice": 0
        }
        # Salva os dados vazios no arquivo JSON
        with open(arquivo_log, 'w', encoding='utf-8') as arq:
            json.dump(dados_vazios, arq, indent=4)

    # Atualiza a planilha de acordo com o que é encontrado nos anexos
    def atualiza_status(self, df: pd.DataFrame, num_pac: str,status: str,save_path: str):
        """
        Updates the 'Status' column and optionally saves to Excel.

        Args:
            df: Input DataFrame
            num_pac: PAC number to search for
            status: 'Sim' or 'Não'
            save_path: If provided, saves the updated DataFrame to this path

        Returns:
            Modified DataFrame
        """
        # Create 'Status' column if missing
        if 'Status' not in df.columns:
            pac_pos = df.columns.get_loc('Nº Reservado PAC') + 1
            df.insert(loc=pac_pos, column='Status', value='')

        # Update status
        mask = df['Nº Reservado PAC'] == num_pac
        if not mask.any():
            print(f"{Fore.RED}⚠️ PAC {num_pac} não encontrado!{Style.RESET_ALL}")
        else:
            df.loc[mask, 'Status'] = status
            df.to_excel(save_path, index=False)
            print(f"{Fore.BLUE}💾 Arquivo salvo em: {save_path}{Style.RESET_ALL}")

    # Acessa uma URL
    def acessa_url(self, url):
        try:
            self.driver.get(url)

        except Exception:
            pass

# Cronometra o tempo de execução da função
def cronometro(func):
        # Decorator que mede o tempo de execução de uma função.'
        @wraps(func)  # Preserva o metadata da função original
        def wrapper(*args, **kwargs):
            start_time = time.time()
            resultado = func(*args, **kwargs)  # Chama a função original
            end_time = time.time()

            # Calcula o tempo que se passou enquanto a função executava
            tempo_transcorrido = end_time - start_time
            horas, rem = divmod(tempo_transcorrido, 3600)
            minutos, segundos = divmod(rem, 60)
            print(f'⏱️ {func.__name__} executada em: {int(horas)}h {int(minutos)}m {segundos:.2f}s')

            return resultado

        return wrapper

@cronometro
def main():
    # Reseta a cor do colorama
    init(autoreset=True)
    # Pasta de download padrão
    pasta_download = r'C:\Users\felipe.rsouza\Downloads'
    # Pasta onde os arquivos serão armazenados
    pasta_destino = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento'
                     r' e Assistência Social\Automações SNEAELIS\Sel PAC DIE')
    # Caminho do arquivo excel
    caminho_arquivo_fonte = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento'
                             r' e Assistência Social\Teste001\tabela_Respostas_Ordenadas.xlsx')
    # Caminho do arquivo que faz o log da execução
    arquivo_log = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento'
                        r' e Assistência Social\Automações SNEAELIS\Sel PAC DIE\progresso.json')

    # [0] Tabela filtrada; [1] Proposta a ser trabalhada; [2] Anexos da Proposta Seleção PAC(tabela);
    # [3] Seleção PAC(volta para filtar por outra proposta
    xpaths = [
              'formListarPropostaPac:dtListarPropostaPac',
              '/html/body/div[3]/div[15]/div[2]/form/table/tbody/tr/td[1]/a',
              'formManterPropostaPac:dt_anexos',
              '//*[@id="breadcrumbs"]/a[2]',
              ]

    # Inicia a classe e estabelece a chamada de funções
    robo = Robo()

    # Extrai dados do Excel: número do processo e tipo de instrumento
    numero_processo, stt = robo.extrair_dados_excel(caminho_arquivo_fonte=caminho_arquivo_fonte,
                                                       busca_id='Nº Reservado PAC',
                                                       status='Status')
    # Marca os links defeituosos para serem pulados
    max_linha = len(numero_processo)
    min_linha = 0

    # Pergunta ao usuário se deseja resetar o log
    reset = input(f'Deseja resetar o log? s/n ')
    if reset.lower() == 's':
        robo.reset(arquivo_log)

    # Carrega o progresso salvo no arquivo de log
    progresso = robo.carrega_progresso(arquivo_log)

    # Define a linha inicial com base no progresso salvo
    if progresso["indice"] > 0:
        min_linha = progresso["indice"]

    # Inicia a sequência de pesquisa
    robo.consulta_instrumento()

    # Loop principal para processar cada linha do Excel
    for indice in tqdm(range(min_linha, max_linha), desc="Scraping...", colour='green', position=0):
        print(Fore.RESET + f'📜 Processo: {numero_processo[indice]}.\n 🏷️ Índice: {indice} de {max_linha}')

        robo.pesquisa_processo(numero_processo=numero_processo[indice])

        try:
            lista_nomes = []
            tabela_proposta = robo.acha_tabela(xpaths[0], mute=True, time=2)
            linha_proposta = robo.linhas(tabela_proposta)

            for nb, linha in enumerate(linha_proposta, start=1):
                if nb > 1:
                    break
                robo.detalhar(linha, xpaths[1], numero_processo[indice])
            try:
                tabela_anexos = robo.acha_tabela(xpaths[2])
                linhas_anexos = robo.linhas(tabela_anexos)
                if not linhas_anexos:
                    print(f'{Fore.RED}❌📊 Nenhuma linha encontrada na tabela{Style.RESET_ALL}')
                    robo.atualiza_status(df=pd.read_excel(caminho_arquivo_fonte),
                                         status='Não',
                                         num_pac=numero_processo[indice],
                                         save_path=caminho_arquivo_fonte
                                         )
                for num, lin in enumerate(linhas_anexos):
                    if num == 0:
                        continue
                    lista_nomes.append(robo.baixa_arquivo(linha=lin, linha_num=num))
                print(lista_nomes)

                robo.atualiza_status(df=pd.read_excel(caminho_arquivo_fonte),
                                     status='Sim',
                                     num_pac=numero_processo[indice],
                                     save_path=caminho_arquivo_fonte
                                     )
            except Exception as e:
                print(e)
            robo.clicar(xpaths[3])

            # Cria pasta de destino para os arquivos a serem transferidos
            pasta_destino_final = robo.criar_pasta(nome_pasta=numero_processo[indice],
                             caminho_pasta_onedrive=pasta_destino,
                                                   )
            # Transfere arquivos baixados
            robo.espera_completar_download(pasta_download=pasta_download)
            time.sleep(0.3)
            robo.transfere_arquivos(caminho_pasta=pasta_destino_final, pasta_download=pasta_download)

            # Salva o progresso do robô
            robo.salva_progresso(arquivo_log=arquivo_log,
                                 processo_visitado=numero_processo[indice],
                                 indice=indice
                                 )
        except TimeoutException:
            print(Fore.RED + f'❌ Erro ao pesquisar o processo {numero_processo[indice]}')
            continue
        except Exception as e:
            print(e)

@cronometro
def captura_tela():
    # Reseta a cor do colorama
    init(autoreset=True)
    # Pasta de download padrão
    pasta_download = r'C:\Users\felipe.rsouza\Downloads'
    # Pasta onde os arquivos serão armazenados
    pasta_destino = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento'
                     r' e Assistência Social\Automações SNEAELIS\Sel PAC DIE')
    # Caminho do arquivo excel
    caminho_arquivo_fonte = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento'
                             r' e Assistência Social\Teste001\tabela_Respostas_Ordenadas.xlsx')
    # Caminho do arquivo que faz o log da execução
    arquivo_log = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento'
                        r' e Assistência Social\Automações SNEAELIS\Sel PAC DIE\progresso_maps.json')

    # Inicia a classe e estabelece a chamada de funções
    robo = Robo()

    # Extrai dados do Excel: número do processo e tipo de instrumento
    numero_processo, status, url = robo.extrair_dados_excel(caminho_arquivo_fonte=caminho_arquivo_fonte,
                                                            busca_id='Nº Reservado PAC',
                                                            status='Status',
                                                            links='Link do Google Maps'
                                                    )
    # Marca os links defeituosos para serem pulados
    max_linha = len(numero_processo)
    min_linha = 0

    # Pergunta ao usuário se deseja resetar o log
    reset = input(f'Deseja resetar o log? s/n ')
    if reset.lower() == 's':
        robo.reset(arquivo_log)

    # Carrega o progresso salvo no arquivo de log
    progresso = robo.carrega_progresso(arquivo_log)

    # Define a linha inicial com base no progresso salvo
    if progresso["indice"] > 0:
        min_linha = progresso["indice"]

    # Loop principal para processar cada linha do Excel
    for indice in tqdm(range(min_linha, max_linha), desc="Scraping...", colour='green'):
        print(Fore.RESET + f'📜 Processo: {numero_processo[indice]}.\n 🏷️ Índice: {indice} de {max_linha}')
        if status[indice].lower() == "não":
            continue
        try:
            robo.driver.get(url[indice])
            robo.webdriver_element_wait('//*[@id="omnibox-singlebox"]/div')
            file_name = numero_processo[indice].split('/')[0] + '.png'
            time.sleep(2)
            print(os.path.join(pasta_download, file_name))
            robo.driver.save_screenshot(os.path.join(pasta_download, file_name))

            # Cria pasta de destino para os arquivos a serem transferidos
            pasta_destino_final = robo.criar_pasta(nome_pasta=numero_processo[indice],
                             caminho_pasta_onedrive=pasta_destino,
                                                      )

            # Transfere arquivos baixados
            robo.espera_completar_download(pasta_download=pasta_download)
            time.sleep(0.1)
            robo.transfere_arquivos(caminho_pasta=pasta_destino_final, pasta_download=pasta_download)

            # Salva o progresso do robô
            robo.salva_progresso(arquivo_log=arquivo_log,
                                 processo_visitado=numero_processo[indice],
                                 indice=indice
                                 )

        except TimeoutException:
            print(Fore.RED + f'❌ Erro ao pesquisar o processo {numero_processo[indice]}')
            continue
        except Exception as e:
            print(e)


if __name__ == "__main__":
    main()
    #captura_tela()