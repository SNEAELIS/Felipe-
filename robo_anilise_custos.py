from selenium import webdriver
from selenium.common import WebDriverException, TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime, timedelta
from selenium.webdriver import ActionChains
from colorama import init, Fore, Back, Style
import pandas as pd
import time
import os
import win32com.client as win32
import sys
import shutil
import json
import traceback

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

    # Navega até a página de busca do instrumento ou proposta
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
                img = reset.find_element(By.XPATH, '//*[@id="logo"]/a/img')
                action = ActionChains(self.driver)
                action.move_to_element(img).perform()
                reset.find_element(By.TAG_NAME, 'a').click()
                print(Fore.MAGENTA + "\n✅ Processo resetado com sucesso !")
        except Exception as e:
            print(Fore.RED + f'🔄❌ Falha ao resetar.')

        # [0] Propostas; [1] Seleção PAC
        xpaths = ["//div[@id='menuPrincipal']/div/div[3]",
                  "//div[@id='contentMenu']//a[normalize-space()='Consultar Propostas']"
                  ]
        try:
            for idx in range(len(xpaths)):
                self.webdriver_element_wait(xpaths[idx]).click()
            print(f"{Fore.MAGENTA}✅ Sucesso em acessar a página de busca de processo{Style.RESET_ALL}")

        except Exception as e:
            print(Fore.RED + f'🔴📄 Instrumento indisponível. \nErro: {e}')
            sys.exit(1)

    # Pesquisa o termo de fomento listado na planilha e executa download e transferência caso exista algúm.
    def loop_de_pesquisa(self, numero_processo: str, caminho_pasta: str, pasta_download: str, feriado: int, err: list=None):
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

        # Baixa os PDF's da tabela HTML
        def baixa_pdf():
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
                time.sleep(0.5)
                tabela = self.driver.find_element(By.ID, 'listaAnexos')

                # Encontra todas as linhas da tabela
                linhas = tabela.find_elements(By.TAG_NAME, 'tr')

                # Diz quantas páginas tem
                paginas = self.conta_paginas(tabela)

                try:
                    for pagina in range(1, paginas + 1):
                        int(pagina)
                        if pagina > 1:
                            self.driver.find_element(By.LINK_TEXT, f'{pagina}').click()
                            # Encontra a tabela na página atual
                            tabela = self.driver.find_element(By.ID, 'listaAnexos')

                            # Encontra todas as linhas da tabela atual
                            linhas = tabela.find_elements(By.TAG_NAME, 'tr')

                        for indice, linha in enumerate(linhas):
                            if indice in err or indice == 0:
                                continue
                            try:
                                # Pega o valor da data em formato string
                                elemento_data_site = linha.find_element(By.CLASS_NAME, 'dataUpload')
                                data_site = elemento_data_site.text if elemento_data_site else None
                                try:
                                    # Compara a data do site com a última data registada na planilha
                                    if self.compara_data(data_site, feriado):
                                        # Pega o elemento do botão de download e deixa selecionado
                                        botao_download = linha.find_element(By.CLASS_NAME, 'buttonLink')
                                        if botao_download:
                                            botao_download.click()
                                except Exception as e:
                                    print(f'❌ Botã de download não encontrado. erro: {type(e).__name__}')
                            except StaleElementReferenceException:
                                try:
                                    linha_erro = indice - 1
                                    print(f"StaleElementReferenceException occurred at line: {linha_erro}")
                                    if linha_erro not in err:
                                        err.append(linha_erro)
                                    self.driver.back()
                                    self.consulta_instrumento()
                                    self.loop_de_pesquisa(numero_processo, caminho_pasta, pasta_download,
                                                      feriado, err)
                                except Exception as error:
                                    error_trace = traceback.format_exc()
                                    print(f'❌ Erro ao pular linha com falha. Erro:'
                                          f' {type(error).__name__}\nTraceback:\n{error_trace}')
                            except Exception as error:
                                print(f"❌ Erro ao processar a linha nº{indice} de termo, erro: {type(error).__name__}")
                                continue
                except Exception as error:
                    print(f"❌ Erro ao buscar nova página: {error}. Err: {type(error).__name__}")
            except Exception as error:
                print(f'❌ Erro de download: Termo {error}. Err: {type(error).__name__}')

        if not err:
            err = []
        try:
            # Seleciona campo de consulta/pesquisa, insere o número de proposta/instrumento e da ENTER
            campo_pesquisa = self.webdriver_element_wait('/html/body/div[3]/div[14]/div[3]/'
                                                         'div/div/form/table/tbody/tr[1]/td[2]/input')
            campo_pesquisa.clear()
            campo_pesquisa.send_keys(numero_processo)
            campo_pesquisa.send_keys(Keys.ENTER)

            # Acessa o item proposta/instrumento
            self.webdriver_element_wait('/html/body/div[3]/div[14]/div[3]/'
                                        'div[3]/table/tbody/tr/td[1]/div/a').click()

            # Seleciona aba primária, após acessar processo/instrumento. Aba Plano de trabalho
            # (aba terciária no total)
            self.webdriver_element_wait("//span[contains(text(),'Plano de Trabalho')]").click()
            # Seleciona aba anexos
            self.webdriver_element_wait("//span[contains(text(),'Anexos')]").click()

            # acessa lista proposta
            self.lista_proposta()

            # Verifica se a tabela existe na página anexos proposta e baixa os anexos.
            try:
                baixa_pdf()

            except Exception:
                print(f"❌ Tabela não encontrada.\nErro:")

            # Seleciona botão voltar
            botao_voltar = self.webdriver_element_wait('/html/body/div[3]/div[14]/div[4]/div/div[1]/'
                                                       'form/table/tbody/tr[1]/td/input')
            botao_voltar.click()

            # Seleciona lista de anexos execução e manda baixar os arquivos
            try:
                botao_lista_execucao = self.webdriver_element_wait('//tbody//tr//input[2]')

                if botao_lista_execucao.is_displayed() or botao_lista_execucao.is_enabled():
                    self.lista_execucao()

                    # Verifica se a tabela existe na página anexos execução.
                    try:
                        baixa_pdf()
                        # Volta para a aba de consulta (começo do loop)
                        consulta_proposta = self.webdriver_element_wait("//a[normalize-space()='Consultar Proposta']")
                        if consulta_proposta:
                            consulta_proposta.click()
                    except Exception as e:
                        print(f"❌ Tabela não encontrada.\nErro: {e}")
            except Exception:
                print(f'❌ Falha em sair da consulta de listas de anexos')
                self.consulta_instrumento()

            # espera os downloads terminarem
            self.espera_completar_download(pasta_download=pasta_download)
            # Transfere os arquivos baixados para a pasta com nome do processo referente
            self.transfere_arquivos(caminho_pasta, pasta_download)

        except TimeoutException:
            self.consulta_instrumento()
            print(f'⏳💀 TIMEOUT')
        except Exception as e:
            print(f'❌ Falha no loop de pesquisa do processo')
            sys.exit(1)

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

            time.sleep(0.1)

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
            paginas = int(paginas[-2])
            return paginas
        except NoSuchElementException:
            paginas = 1
            return paginas

    # Pega a data do documento e o nome da aba
    def compara_data(self, data_site: str, feriado: int) -> bool:
        """
        Compara a data fornecida com a data de ontem.

        Args:
            data_site: Data a ser comparada, no formato 'AAAA-MM-DD' ou 'AAAA-MM-DD HH:MM:SS'.
            feriado: Quantidade de dias a serem descontadas para pegar os dias que o robo não rodou

        Returns:
            True se a data fornecida for anterior à data de ontem, False caso contrário.
            Retorna False se o formato da data fornecida for inválido.
        """
        # Obtém a data de ontem
        try:
            return True
            # Converte a data do site para datetime, tratando diferentes formatos
            data_site_dt = self.converter_data(data_site)
            if data_site_dt is None:
                return False  # Formato de data inválido
            if datetime.isoweekday(datetime.today()) == 1:
                data_ontem = datetime.now() - timedelta(days=2 + feriado)
            else:
                data_ontem = datetime.now() - timedelta(days=feriado)

            # Comparação
            if data_site_dt >= data_ontem:
                print(data_site_dt)
                print(data_ontem)
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

        print(f"⚠️📆 Formato de data inválido: {data}")
        return None

    # Acessa a lista de anexos proposta
    def lista_proposta(self) -> None:
        """
               Acessa a lista de anexos da proposta.

               Esta função clica no botão para exibir a lista de anexos da proposta
               e define o nome da coluna que contém as datas dos anexos.

               Returns:
                   None.
               """
        # Seleciona lista de anexos proposta e acessa a mesma
        try:
            lista_anexos_proposta = self.webdriver_element_wait('/html/body/div[3]/div[14]/div[3]/div[1]/div/'
                                                                'form/table/tbody/tr/td[2]/input[1]')
            lista_anexos_proposta.click()
        except Exception:
            print("🤷‍♂️❌ Erro ao tentar entar na lista de anexos propostas")

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
            print(f"🤷‍♂️❌ Erro ao tentar entar na lista de anexos execução: {e}")

    # Baixa os PDF's da tabela HTML

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
                        print(f'📂 Movimentado arquivo: {arq}')

            except Exception as e:
                print(f'❌ Falha ao mover arquivo {e}')

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
            exist_ok = True  # evita erro se a pasta já existir.
            os.makedirs(caminho_pasta, exist_ok=True)
            print(f"✅ Pasta '{nome_pasta}' criada em: {caminho_pasta}\n")
        except Exception as e:
            print(f"❌ Erro ao criar a pasta '{nome_pasta}': {e}")
        # Retorna o caminho completo da pasta, mesmo que a criação tenha falhado (para tratamento posterior)
        return caminho_pasta

    # Função para enviar e-mails com anexos.
    def enviar_email_tecnico(self, email_destino: str, destinatario: str,
                             lista_documentos: list or dict, email_copia: any = False, mensagem=False,
                             numero_updates=0) -> None:
        """
                Envia um e-mail técnico para o destinatário especificado.

                Esta função envia um e-mail para o técnico, informando sobre a atualização da proposta
                e fornecendo o número do processo e o caminho para a pasta onde o documento atualizado
                está localizado.

                Args:
                :param    email_destino: O endereço de e-mail do destinatário.
                :param    destinatario: O nome do destinatário.
                :param    lista_documentos: O número do processo está na posição 0, o caminho da pasta
                          está na posição 1 e a posição 2 tem a lista de todos os
                          documentos que foram atualizados daquele processo.
                :param    mensagem: Define qual mensagem vai ser enviada, se é para o técnico ou para os chefes.

                Returns:
                    None
                """

        # Lista_documentos recebe argumentos da variável docs_atuais
        link_onedrive = (r'https://mdsgov-my.sharepoint.com/:f:/g/personal/'
                         r'andrei_rodrigues_esporte_gov_br/Eu5WAkT4dFlItHjvxYADDLEBJ2JfPnGcZTt6xiKGj1AMjw?e=03H12b')
        # Cria o corpo do e-mail com as informações necessárias para os Chefes
        if mensagem:
            mensagem = f"""

            <p>Prezados Sr. Paulo Afonso de Araujo Quermes, Leidiane Rodrigues Pires e Carla Prado Novais</p>

            <p>Gostaria de informar que a proposta houveram {numero_updates} atualizações.</p>

            <p>Segue a lista dos documentos:<br><br> {lista_documentos}. <br><br>Link da pasta <a href="{link_onedrive}">{'Resultados Robô Mérito-Custos'}</a></p>

            <p> Atenciosamente,</p>
        """
            try:
                copia_list = [email_destino, email_copia, destinatario]
                # Cria integração com o outlook
                outlook = win32.Dispatch('outlook.application')

                # Configurar e-mail
                email = outlook.CreateItem(0)
                email.To = f'{';'.join(copia_list)}'
                email.Subject = f'Atualização Diária das propostas'
                email.HTMLBody = mensagem

                email.Send()
                print(f"📧📤 E-mail enviado para {copia_list}")
            except Exception as e:
                print(f"❌ Falha ao enviar e-mail chefes: \n{e}")
        # Cria o corpo do e-mail com as informações necessárias para cada técnico
        else:
            linha_formatada = '<br>'.join(lista_documentos)

            # Cria o corpo do e-mail com as informações necessárias para os técnicos
            mensagem = f"""

            <p>Prezado(a) {destinatario.capitalize()}</p>

            <p>Gostaria de informar as atualizações diárias para as propostas de seu encargo.</p>

            <p>O documento baixado se encontra na pasta <a href="{link_onedrive}">{'Resultados Robô Mérito-Custos'}</a>.</p>


            <p>Segue a lista das propostas:<br> {linha_formatada}.</p>

            <p> Atenciosamente</p>

            <p></p>
            <p></p>
            <p></p>
            <p>ATENÇÂO! ESTAS PROPOSTAS SÃO MONITORADAS NO TRANSFEREGOV.COM</p>
        """

            try:
                # Cria integração com o outlook
                outlook = win32.Dispatch('outlook.application')

                # Configurar e-mail
                email = outlook.CreateItem(0)
                email.To = f'{email_destino}'
                email.Subject = f'Atualização de Propostas'
                email.HTMLBody = mensagem
                email.Send()
                print(f"✅ E-mail enviado para {destinatario}, no endereço {email_destino}")

            except Exception:
                print(f"❌ Falha ao enviar e-mail para: {destinatario}\n No e-mail: {email_destino}")

    def extrair_dados_excel(self, caminho_arquivo_fonte, busca_id: str, email_id: str,
                            nome_recipiente_id: str, tipo_instrumento_id: str, email_copia_id: str):
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
        dados_processo.replace(u'\xa0', '', regex=True).infer_objects(copy=True)

        # Cria um lista para cada coluna do arquivo xlsx
        numero_processo = list()
        email = list()
        entidade = list()
        tipo_instrumento = list()
        email_copia = list()

        try:
            # Itera a planilha e armazena os dados em listas
            for indice, linha in dados_processo.iterrows():  # Assume que a primeira linha e um cabeçalho
                numero_processo.append(linha[busca_id])  # Busca o número do processo
                email.append(linha[email_id])  # Busca o destinatário da mensagem
                entidade.append(linha[nome_recipiente_id])  # Busca destinatário do email
                tipo_instrumento.append(linha[tipo_instrumento_id])  # Busca o tipo de instrumento
                email_copia.append(linha[email_copia_id])  # Busca o endereço de email para enviar cópia
        except Exception as e:
            print(f"❌ Erro de leitura encontrado, erro: {e}")

        return numero_processo, email, entidade, tipo_instrumento, email_copia

    def login(self, login: str, senha: str):
        try:
            # Acessa a tela de login do gov.br
            pre_login = self.webdriver_element_wait('/html[1]/body[1]/div[2]/div[1]/div[1]/div[3]/main[1]/form[1]/a[1]')
            pre_login.click()
            # Entra o cpf od usuário
            login_gov = self.webdriver_element_wait('/html[1]/body[1]/div[1]/main[1]/form[1]/div[1]/div[2]/input[1]')
            login_gov.clear()
            login_gov.send_keys(login)
            login_gov.send_keys(Keys.ENTER)
            # Entra a senha do usuário
            entra_senha = self.webdriver_element_wait('/html[1]/body[1]/div[1]/main[1]/form[1]/div[1]/div[1]/input[1]')
            entra_senha.clear()
            entra_senha.send_keys(senha)
            entra_senha.send_keys(Keys.ENTER)
        except Exception as e:
            print(f"❌ Erro de login encontrado, erro: {e}")

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
    def salva_progresso(self, arquivo_log: str, processo_visitado: str, arquivos_baixados: list, indice: int):
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
            print(f"⚠️Nenhum documento novo encontrado para o processo {numero_processo}.\n")
            return []

    def timmer(self, tm=2):
        return time.sleep(tm)


def main(feriado=2) -> None:
    # CASO TENHA QUE ITERAR TODAS AS PASTAS E BAIXAR TODOS OS ARQUIVOS NOVAMENTE
    # lista_pagina_quebrada = [61, 352, 482,604]
    # Caminho da pasta download que é o diretório padrão para download. Use o caminho da pasta 'Download' do
    # seu computador
    pasta_download = r'C:\Users\felipe.rsouza\Downloads'

    # Caminho do arquivo .xlsx que contem os dados necessários para rodar o robô
    caminho_arquivo_fonte = (r'C:\Users\felipe.rsouza\Documents\Dataframe Custos e Méritos.xlsx')

    # Rota da pasta onde os arquivos baixados serão alocados, cada processo terá uma subpasta dentro desta
    caminho_pasta_onedrive = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e'
                              r' Assistência Social\Automações SNEAELIS\Resultados Robô Mérito-Custos')

    # Caminho do arquivo JSON que serve como catão de memória
    arquivo_log = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento'
                   r' e Assistência Social\Automações SNEAELIS\Resultados Robô Mérito-Custos( back_end )\arquivo_log.json')

    # Instancia um objeto da classe Robo
    robo = Robo()
    # Extrai dados de colunas específicas do Excel
    numero_processo, email, entidade, tipo_instrumento, email_copia = robo.extrair_dados_excel(
        caminho_arquivo_fonte=caminho_arquivo_fonte, busca_id='Num Proposta', email_id='E-mail Técnico Custos',
        nome_recipiente_id='Responsável  Análise Custos', tipo_instrumento_id='Tipo Instrumento:',
        email_copia_id='Técnicos Ludmila')
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
        print(f"{indice},   >>>  {(indice / max_linha) * 100:.2f}%")
        # robo.timmer(3)
        if tipo_instrumento[indice] in ['Termo de Fomento', 'Termo de Colaboração']:
            # Cria pasta com número do processo
            caminho_pasta = robo.criar_pasta(nome_pasta=numero_processo[indice],
                                             caminho_pasta_onedrive=caminho_pasta_onedrive,
                                             tipo_instrumento=tipo_instrumento[indice])
            # Executa pesquisa dos termos e salva os resultados na pasta "caminho_pasta"
            robo.loop_de_pesquisa(numero_processo[indice], caminho_pasta, pasta_download, feriado)

            # Confirma se houve atualização na pasta e envia email para o técnico
            confirma_email = list(robo.condicao_email(numero_processo=numero_processo[indice],
                                                      caminho_pasta=caminho_pasta))
            if confirma_email:
                robo.salva_progresso(arquivo_log, numero_processo[indice], confirma_email[2], indice=indice)
        else:
            # Cria pasta com número do processo
            caminho_pasta = robo.criar_pasta(numero_processo[indice],
                                             caminho_pasta_onedrive=caminho_pasta_onedrive,
                                             tipo_instrumento=tipo_instrumento[indice])
            # Executa pesquisa dos termos e salva os resultados na pasta "caminho_pasta"
            robo.loop_de_pesquisa(numero_processo[indice], caminho_pasta, pasta_download, feriado)

            # Confirma se houve atualização na pasta e envia email para o técnico
            confirma_email = list(robo.condicao_email(
                numero_processo=numero_processo[indice], caminho_pasta=caminho_pasta)
            )
            if confirma_email:
                robo.salva_progresso(arquivo_log=arquivo_log, processo_visitado=numero_processo[indice], arquivos_baixados=confirma_email[2], indice=indice)


def count_empty_folders(directory):
    """Finds and prints the names of empty folders within a given directory."""
    empty_folders = []
    for item in os.listdir(directory):
        item_path = os.path.join(directory, item)
        if os.path.isdir(item_path):
            if not os.listdir(item_path):
                empty_folders.append(item)
    if empty_folders:
        print("Empty folders found:")
        for folder_name in empty_folders:
            parts = folder_name.split('_')
            parts = '/'.join(parts[:2])
            print(folder_name)
            print(parts)
    else:
        print(f"No empty folders found in '{directory}'.")


if __name__ == "__main__":
    start_time = time.time()

    #main()
    count_empty_folders(r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e '
                        r'Assistência Social\Automações SNEAELIS\Resultados Robô Mérito-Custos')

    end_time = time.time()
    tempo_total = end_time - start_time
    horas = int(tempo_total // 3600)
    minutos = int((tempo_total % 3600) // 60)
    segundos = int(tempo_total % 60)
    print(f'⏳ Tempo de execução: {horas}h {minutos}m {segundos}s')