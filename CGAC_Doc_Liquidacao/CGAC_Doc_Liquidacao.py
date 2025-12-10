from selenium import webdriver
from selenium.common import WebDriverException, TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
import sys
import pandas as pd
import time
import os
import csv
import shutil
import json


class Robo:
    def __init__(self):
        """
        Inicializa o objeto Robo, configurando e iniciando o driver do Chrome.
        """
        try:
            # Contador para o número de links encontrados
            self.contador = 0

            # Configuração do registro
            # Inicia as opções do Chrome
            self.chrome_options = webdriver.ChromeOptions()
            # Endereço de depuração para conexão com o Chrome
            self.chrome_options.add_argument("--disable-extensions")  # Disables all extensions
            self.chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
            # Inicializa o driver do Chrome com as opções e o gerenciador de drivers
            self.driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()),
                options=self.chrome_options)

            qnt_abas = self.driver.window_handles
            for handle in qnt_abas:
                self.driver.switch_to.window(handle)
                url = self.driver.current_url
                if "chrome" in url:
                    qnt_abas.remove(handle)

            self.driver.switch_to.window(qnt_abas[0])

            print("✅ Conectado ao navegador existente com sucesso.", "\nCurrent URL:",
                  self.driver.current_url)

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
            arq_crdownload = [f for f in os.listdir(pasta_download) if f.endswith('.crdownload')]
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
            - Se o arquivo foi modificado **hoje ou depois**, move-o para `caminho_pasta`.
            - Exibe uma mensagem no console para cada arquivo movido.
            """
        # pega a data do dia que está executando o programa
        data_hoje = self.data_hoje()
        tempo_final = time.time() + tempo_limite
        self.permuta_nome_arq(pasta_download)

        counter = 0
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
                        counter += 1
                if not os.listdir(pasta_download):
                    break
            except Exception as e:
                print(f'❌ Falha ao mover arquivo {e}')

        print(f'📂🚚 Arquivos Movimentados. Total de arquivos: {counter}')

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

    def extrair_dados_excel(self, caminho_arquivo_fonte, busca_id: str, tipo_instrumento_id: str):
        """
           Lê os contatos de uma planilha Excel e executa ações baseadas nos dados extraídos.

           Args:
               caminho_arquivo_fonte (str): Caminho do arquivo Excel que será lido.
               busca_id (str): Nome da coluna que contém os números de processo.
               email_id (str): Nome da coluna que contém os endereços de e-mail.
               nome_recipiente_id (str): Nome da coluna que contém o nome da entidade ou destinatário.
               tipo_instrumento_id (str): Nome da coluna que contém o tipo de processo.
           """
        dados_processo = pd.read_excel(caminho_arquivo_fonte, dtype=str)
        dados_processo = dados_processo.replace(u'\xa0', '', regex=True)
        dados_processo = dados_processo.infer_objects(copy=False)

        # Cria um lista para cada coluna do arquivo xlsx
        numero_processo = list()
        tipo_instrumento = list()
        try:
            # Itera a planilha e armazena os dados em listas
            for indice, linha in dados_processo.iterrows():  # Assume que a primeira linha e um cabeçalho
                numero_processo.append(linha[busca_id])  # Busca o número do processo
                tipo_instrumento.append(linha[tipo_instrumento_id]) # Busca o tipo de instrumento

        except Exception as e:
            print(f"❌ Erro de leitura encontrado, erro: {e}")

        return numero_processo, tipo_instrumento

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
        return WebDriverWait(self.driver, 3).until(EC.element_to_be_clickable((By.XPATH, xpath)))

    # Consulta o instrumento ou proposta (faz todo o caminho até chegar no item desejado)
    def consulta_instrumento(self):
        """
               Navega pelas abas do sistema até a página de busca de processos.

               Esta função clica nas abas principal e secundária para acessar a página
               onde é possível realizar a busca de processos.
               """
        try:
            # Reseta para página inicial
            time.sleep(0.3)
            reset = self.driver.find_element(By.XPATH, '//*[@id="logo"]/a/img')
            if reset:
                reset = self.driver.find_element(By.XPATH, '/html[1]/body[1]/div[3]/div[2]/div[1]/a[1]/img[1]')
                reset.click()
                print(f"✅ Processo resetado com sucesso !")
        except Exception as e:
            print(f'❌ Falha ao resetar, erro:')
            pass
        try:
            # Seleciona a aba principal desejada. Aba Execução
            aba_principal = self.webdriver_element_wait("//div[normalize-space()='Execução']")
            aba_principal.click()

            # Seleciona a aba secundária desejada. Sub-Aba Consultar pré-instrument
            aba_secundaria = self.webdriver_element_wait("//div[@id='contentMenu']//"
                                                         "div[1]//ul[1]//li[6]//a[1]")
            aba_secundaria.click()
            print(f"✅ Sucesso em acessar a página de busca de processo")

        except Exception as e:
            print(f'❌ Falha ao consultar o processo {str(e)[:80]}')
            sys.exit(1)

    # Loop que acessa os vários instrumentos para buscar os dados
    def loop_instrumentos(self, numero_processo):
        # Seleciona campo de consulta/pesquisa, insere o número de proposta/instrumento e da ENTER
        try:
            campo_pesquisa = self.webdriver_element_wait("//input[@id='consultarNumeroConvenio']")
            campo_pesquisa.clear()
            campo_pesquisa.send_keys(numero_processo)
            campo_pesquisa.send_keys(Keys.ENTER)

            # Acessa o item proposta/instrumento
            self.webdriver_element_wait('//*[@id="instrumentoId"]/a').click()

            # Seleciona aba primária, após acessar processo/instrumento. Aba Execução Covenente
            self.webdriver_element_wait("//span[contains(text(),'Execução Convenente')]").click()
            # Seleciona aba Documentos de Liquidação
            self.webdriver_element_wait("//span[contains(text(),'Documento de Liquidação')]").click()
            return False
        except TimeoutException:
            print(f'🔴📄 Instrumento indisponível')
            return True
    # Conta o número de páginas de tabelas o instrumento tem
    def conta_paginas(self, tabela) -> int:
        # Diz quantas páginas tem
        try:
            paginas = tabela.find_element(By.TAG_NAME, 'span').text
            paginas = paginas.split('(')[0]
            paginas.strip()
            paginas = int(paginas[-3: -1])
            return paginas
        except NoSuchElementException:
            paginas = 1
            return paginas

    def volta_pesquisa(self):
        try:
            self.webdriver_element_wait('//*[@id="breadcrumbs"]/a[2]').click()
        except NoSuchElementException as nse:
            print(f'👀🚫 {nse}')
        except Exception as e:
            print(f'❌ {e}')

    # Abre todos os links da tabela em uma nova aba
    def abre_abas(self, linhas):
        """
        Abre todos os links de uma tabela em novas abas.
        """
        # Pula a primeira linha (cabeçalho)
        links = linhas[1:]
        try:
            for link in links:
                # Abre cada link em uma nova aba usando Ctrl + Clique
                abre_aba = link.find_element(By.CSS_SELECTOR, "td a[href]")
                webdriver.ActionChains(self.driver).key_down(Keys.CONTROL).click(abre_aba).key_up(
                    Keys.CONTROL).perform()
            print(f"✅ {len(links)} abas abertas com sucesso.")
            abas = self.driver.window_handles
            for handle in abas:
                self.driver.switch_to.window(handle)
                url = self.driver.current_url
                if "chrome" in url:
                    abas.remove(handle)

            return int(len(abas))
        except Exception:
            print(f"❌ Erro ao abrir abas ")

    # Muda para a aba de índice dado pelo argumento num_aba
    def muda_aba(self, num_aba: int):
        """
        Muda o foco para a aba especificada pelo índice.
        """
        # Muda para a aba de índice [1]
        time.sleep(0.3)
        try:
            qnt_abas = self.driver.window_handles
            for handle in qnt_abas:
                self.driver.switch_to.window(handle)
                url = self.driver.current_url
                if "chrome" in url:
                    qnt_abas.remove(handle)

            self.driver.switch_to.window(qnt_abas[num_aba])
        except Exception as e:
            print(f"❌ Falha ao tentar mudar de aba.\nError: {type(e).__name__}.\nFalha: {str(e)[:80]}")

    # fecha a aba atual
    def fecha_aba(self):
        """
        Fecha a aba especificada pelo índice e muda o foco para a aba anterior.
        """
        try:
            self.driver.close()
        except Exception:
            print(f"❌ Não consegui fechar a aba.")

    def proxima_pag(self, pagina: int, max_pagina: int):
        try:
            if pagina < max_pagina:
                pagina += 1
                time.sleep(0.3)
                self.driver.find_element(By.LINK_TEXT, f'{pagina}').click()
                print(f"➡️📄 Mudando para página {pagina}/{max_pagina}")
        except Exception as e:
            print(f'📄❌ Erro ao passar de página {e}')

    def acha_tabela(self, nome_tabela: str):
        try:
            # Encontra a tabela na página atual
            tabela = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.ID, nome_tabela)))
            return tabela
        except Exception as e:
            print(f"❌ Erro ao buscar itens da tabela: {type(e).__name__}")

    # Acha as linhas da tabela
    def linhas(self, tabela):
        try:
            linhas = tabela.find_elements(By.TAG_NAME, 'tr')
            return linhas
        except Exception as e:
            print(f'❌ Falha ao tentar achar as linhas: {type(e).__name__}')

    # Baixa os arquivos
    def baixa_arquivo(self, linhas):
        # Separa as linhas da tabela para acessar os links individuais
        self.conta_arquivos(linhas=linhas)

        for linha in linhas[1: ]:
            try:
                # Pega o elemento do botão de download e deixa selecionado
                botao_download = WebDriverWait(linha, 3).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, 'buttonLink')))
                if botao_download:
                    continue
                    botao_download.click()
            except:
                print('❌ Botã de download não encontrado.')

    def conta_arquivos(self, linhas):
        for linha in linhas[1: ]:
            try:
                # Pega o elemento do botão de download e deixa selecionado
                botao_download = WebDriverWait(linha, 10).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, 'buttonLink')))
                if botao_download:
                    self.contador += 1
            except:
                print('❌ Botã de download não encontrado.')

    # Salva o progresso em um arquivo json
    def salva_progresso(self, arquivo_log: str, processo_visitado: str, indice: int):
        """
        Salva o progresso atual em um arquivo JSON.

        :param arquivo_log: Endereço do arquivo JSON
        :param processo_visitado: Lista de processos concluídos.
        :param arquivos_baixados: Lista de arquivos baixados.
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
            "processo_visitado": 000000,
            "indice": 0
        }
        # Salva os dados vazios no arquivo JSON
        with open(arquivo_log, 'w', encoding='utf-8') as arq:
            json.dump(dados_vazios, arq, indent=4)

    # Salva quantidade de itens encontrados em arquivo csv.
    def salva_qtd(self, file_path, draft_num: str,):
        file_exists = os.path.exists(file_path)

        with open(file_path, 'a', newline='', encoding='utf-8') as f:
            writer = csv.writer(f)

            if not file_exists:
                writer.writerow(['Proposta', 'Qtd_Itens'])

            writer.writerow([draft_num, self.contador])


def conta_arquivos(caminho_pasta):
    return sum(1 for item in os.listdir(caminho_pasta)
               if os.path.isfile(os.path.join(caminho_pasta, item)))

def main():
    pasta_download = r'C:\Users\felipe.rsouza\Downloads'

    pasta_destino = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e'
                              r' Assistência Social\Automações SNEAELIS\CGAC_2024')

    caminho_arquivo_fonte = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e '
                             r'Assistência Social\Automações '
                             r'SNEAELIS\CGAC_2024\CGAC_2024_dataSource_filtered.xlsx')

    cgac_arquivo_log = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e '
                        r'Assistência Social\Automações SNEAELIS\CGAC_2024\CGAC_arquivo_log.json')

    arq_contador = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Automações '
                    r'SNEAELIS\CGAC_2024\conta_arquivos.csv')


    # XPaths para navegação e limpeza de campos
                        #  Próxima página                         Limpar Campos
    loop_saida_bug = ["//a[normalize-space()='Próx']", "//td[@class='FormLinhaBotoes']//input[2]"]

    total_file = dict()

    robo = Robo()

    # Extrai dados do Excel: número do processo e tipo de instrumento
    numero_processo, tipo_instrumento= robo.extrair_dados_excel(caminho_arquivo_fonte=caminho_arquivo_fonte,
                                                                busca_id='Nº CONVÊNIO',
                                                                tipo_instrumento_id='MODO')

    max_linha = len(numero_processo)

    # Pergunta ao usuário se deseja resetar o log
    reset = input(f'Deseja resetar o log? s/n ')
    if reset == 's':
        robo.reset(cgac_arquivo_log)

    # Carrega o progresso salvo no arquivo de log
    progresso = robo.carrega_progresso(cgac_arquivo_log)

    # Define a linha inicial com base no progresso salvo
    if progresso["indice"] > 0:
        min_linha = progresso["indice"]
    else:
        min_linha = 0

    rmk = ['941070', '941666', '941407', '888699', '882017', '897333', '940697', '909989', '909765', '913320', '928420', '925902', '925901', '897751', '954489', '740366', '897676', '935351', '943210', '897785', '936889', '922808', '941154', '905377', '928318', '898083', '917388', '909986', '930085', '940945', '920631', '924083', '934702', '910735', '898406', '881440', '942809']



    # Loop principal para processar cada linha do Excel
    for indice in range(min_linha, max_linha):
        print(f'Processo: {numero_processo[indice]}.\n Índice: {indice} do total: {max_linha-1}')

        if numero_processo[indice] in ['878411', '881200', '898261', '905061', '936600']:
            continue
        elif numero_processo[indice] not in rmk:
            continue

        # Salva o progresso atual no arquivo de log
        robo.salva_progresso(cgac_arquivo_log, numero_processo[indice], indice)
        robo.consulta_instrumento()
        loop = robo.loop_instrumentos(numero_processo[indice])

        # Salva o número do processo que causou erro no loop
        if loop:
            with open(r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                      r'Social\Power BI\Python\Requerimento de serviço\CGAG(Cássia)\log_erros.txt', 'a'
              ) as file:
                file.write(f'{numero_processo[indice]}\n')
            continue
        tabela = robo.acha_tabela(nome_tabela='notasFiscais')
        paginas = robo.conta_paginas(tabela=tabela)
        incremento_pagina = 0


        # Loop para navegar pelas páginas da tabela
        for pagina in range(1, paginas+1):
            tabela = robo.acha_tabela(nome_tabela='notasFiscais')
            try:
                # Lógica para páginas após a 10ª e que não são múltiplas de 10
                if 10 < pagina and pagina % 10 != 0 :
                    linhas = robo.linhas(tabela=tabela)
                    abas = robo.abre_abas(linhas=linhas)
                    time.sleep(1)

                    # Loop para processar cada aba aberta
                    while abas > 1:
                        robo.muda_aba(-1)
                        tabela = robo.acha_tabela(nome_tabela='arquivos')
                        linhas = robo.linhas(tabela=tabela)
                        robo.baixa_arquivo(linhas=linhas)
                        robo.fecha_aba()
                        abas -= 1

                    robo.muda_aba(0)

                    if not pagina == paginas:
                        robo.proxima_pag(pagina=pagina, max_pagina=paginas)

                        # Limpa campos e navega para a próxima página
                        WebDriverWait(robo.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, loop_saida_bug[1]))).click()

                        # Aperta o [PROX] a quantidade de vezes necessária
                        for inc in range(incremento_pagina):
                            WebDriverWait(robo.driver, 10).until(
                                EC.element_to_be_clickable((By.XPATH, loop_saida_bug[0]))).click()

                        robo.proxima_pag(pagina=pagina, max_pagina=paginas)

                # Lógica para páginas múltiplas de 10
                elif pagina % 10 == 0:
                    incremento_pagina += 1
                    print(f'Pagina {pagina} aberta')
                    linhas = robo.linhas(tabela=tabela)
                    abas = robo.abre_abas(linhas=linhas)
                    time.sleep(1)

                    while abas > 1:
                        robo.muda_aba(-1)
                        tabela = robo.acha_tabela(nome_tabela='arquivos')
                        linhas = robo.linhas(tabela=tabela)
                        robo.baixa_arquivo(linhas=linhas)
                        robo.fecha_aba()
                        abas -= 1

                    robo.muda_aba(0)

                    # Clica no link de próxima aba
                    WebDriverWait(robo.driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                    loop_saida_bug[0]))).click()

                    # Limpa campos e navega para a próxima página
                    WebDriverWait(robo.driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, loop_saida_bug[1]))).click()

                    # Aperta o [PROX] a quantidade de vezes necessária
                    for inc in range(incremento_pagina):
                        WebDriverWait(robo.driver, 10).until(
                            EC.element_to_be_clickable((By.XPATH, loop_saida_bug[0]))).click()

                # Lógica para as primeiras 10 páginas
                else:
                    linhas = robo.linhas(tabela=tabela)
                    abas = robo.abre_abas(linhas=linhas)
                    time.sleep(1)

                    while abas > 1:
                        robo.muda_aba(-1)
                        tabela = robo.acha_tabela(nome_tabela='arquivos')
                        linhas = robo.linhas(tabela=tabela)
                        robo.baixa_arquivo(linhas=linhas)
                        robo.fecha_aba()
                        abas -= 1

                    robo.muda_aba(0)
                    robo.proxima_pag(pagina=pagina, max_pagina=paginas)
                    robo.driver.find_element(By.XPATH, loop_saida_bug[1]).click()
                    robo.proxima_pag(pagina=pagina, max_pagina=paginas)

            except Exception:
                pass

        pasta_destino_final = robo.criar_pasta(nome_pasta=numero_processo[indice],
                         caminho_pasta_onedrive=pasta_destino,
                         tipo_instrumento=tipo_instrumento[indice])

        # Cria pasta de destino e transfere arquivos baixados
        robo.espera_completar_download(pasta_download=pasta_download)
        robo.transfere_arquivos(caminho_pasta=pasta_destino_final, pasta_download=pasta_download)
        print(f'Número de arquivos na pasta final: {conta_arquivos(caminho_pasta=pasta_destino_final)}.\nNúmero de '
              f'links encontrados: {robo.contador} ')
        total_file[numero_processo[indice]] = robo.contador

        robo.salva_qtd(file_path=arq_contador, draft_num=numero_processo[indice])
        robo.contador = 0

    return total_file


def cont_files_dir():
    root_dir = (
        r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Automações SNEAELIS\CGAC_2024')
    files_per_dir = dict()

    for root, dirs, files in os.walk(root_dir):
        for d in dirs:
            try:
                base_name = d.split('_')[0]
            except IndexError:
                # Handle directory names without an underscore
                base_name = d

            try:
                cnt = 0

                for f in os.listdir(os.path.join(root, d)):
                    f_path = os.path.join(root, d, f)
                    if os.path.isfile(f_path):
                        cnt += 1
            except:
                print(f"Warning: Could not access directory {os.listdir(os.path.join(root, d))} due to permission "
                      f"error. Skipping.")
                continue  # Skip to the next directory

            if base_name not in files_per_dir:
                files_per_dir[base_name] = cnt

    return files_per_dir


def compare_dict(dict_1: dict, dict_2: dict):
    def salve_to_txt(err_keys, file_path):
        line = f'{err_keys}\n'
        with open(file_path, 'a', encoding='utf-8') as f:
            f.write(line)

    file_path = r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Automações SNEAELIS\CGAC_2024\diff_.txt'
    err_keys = list()

    for key in dict_1:
        if key in dict_2:
            if dict_1[key] != dict_2[key]:
                err_keys.append(key)
        else:
            print(f"Key '{key}': DISPARITY FOUND! (Key missing in Dict2)")
    salve_to_txt(err_keys=err_keys, file_path=file_path)
    print(err_keys)


if __name__ == "__main__":
    start_time = time.time()

    total_file_site = main()
    total_file_dir = cont_files_dir()

    compare_dict(total_file_site, total_file_dir)

    end_time = time.time()
    tempo_total = end_time - start_time
    horas = int(tempo_total// 3600)
    minutos = int((tempo_total % 3600) // 60)
    segundos = int(tempo_total % 60)
    print(f'⏳ Tempo de execução: {horas}h {minutos}m {segundos}s')

