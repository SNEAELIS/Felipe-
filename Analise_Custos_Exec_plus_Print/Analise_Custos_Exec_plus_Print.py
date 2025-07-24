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
from pathlib import Path
import pandas as pd
import time
import os
import sys
import shutil
import json
import traceback
import fitz  # PyMuPDF
import hashlib
import pyautogui
import re
import zipfile




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
            print(Fore.RED + f'🔄❌ Falha ao resetar.\nErro: {type(e).__name__}\n{str(e)[:50]}')

        # [0] Excução; [1] Consultar Pré-Instrumento/Instrumento
        xpaths = ['//*[@id="menuPrincipal"]/div[1]/div[4]',
                  '//*[@id="contentMenu"]/div[1]/ul/li[6]/a'
                  ]
        try:
            for idx in range(len(xpaths)):
                self.webdriver_element_wait(xpaths[idx]).click()
            print(f"{Fore.MAGENTA}✅ Sucesso em acessar a página de busca de processo{Style.RESET_ALL}")
        except Exception as e:
            print(Fore.RED + f'🔴📄 Instrumento indisponível. \nErro: {e}')
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
            print(f' Falha ao inserir número de processo no campo de pesquisa. Erro: {type(e).__name__}')

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

    def busca_propostas(self):
        try:
            # Executa pesquisa de anexos
            print('🔁📎 Executando pesquisa de anexos')

            # Aba Plano de trabalho
            self.webdriver_element_wait('//*[@id="div_997366806"]').click()

            # Seleciona aba anexos
            self.webdriver_element_wait('//*[@id="menu_link_997366806_1965609892"]/div').click()

        except Exception as e:
            print(f"Ocorreu um erro ao executar ao pesquisar anexos: {type(e).__name__}"
                  f".\n Erro {e[:50]}")

    # Pesquisa o termo de fomento listado na planilha e executa download e transferência caso exista algúm.
    def loop_de_pesquisa(self, numero_processo: str, caminho_pasta: str, pasta_download: str,
                          feriado: int, err: list=None,
                         pg: int=0, anexos: bool=True):
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
                        print(f"⚠️ Erro ao deletar {file}: {e}")

            print(f"✅ Arquivo ZIP criado em: {zip_path}")
            return str(zip_path)

        # Baixa os PDF's da tabela HTML
        def baixa_pdf_exec():
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

                try:
                    for pagina in range(1, paginas + 1):
                        int(pagina)
                        if pg > pagina:
                            print('🌀📄 pulando página')
                            continue

                        # Ensure the error log has a list for the current page
                        while len(err) < pagina:
                            err.append([])

                        # Check if the last page was the tenth page, this site has a max of 10 pages /block
                        if pagina > 1:
                            if (pagina - 1) % 10 == 0:
                                self.driver.find_element(By.XPATH, '//*[@id="listaAnexos"]/span[2]/a[10]').click()

                            elif (pagina - 1) % 10 != 0:
                                self.driver.find_element(By.LINK_TEXT, f'{pagina}').click()

                        # Encontra a tabela na página atual
                        tabela = self.driver.find_element(By.ID, 'listaAnexos')
                        # Encontra todas as linhas da tabela atual
                        linhas = tabela.find_elements(By.XPATH, './/tbody/tr')

                        for indice, linha in enumerate(linhas):
                            # Pagina - 1 is to correct the index, synce pagina starts @ 1 and list idx @ 0
                            if err[pagina-1]:
                                if indice <= err[pagina - 1][-1]:
                                    print(f'⏭️ pulando linha: {indice}')
                                    continue

                            try:
                                botao_download = linha.find_element(By.CLASS_NAME, 'buttonLink')
                                if botao_download:
                                    botao_download.click()
                            except StaleElementReferenceException:
                                try:
                                    linha_erro = indice - 1
                                    print(f"⚠️ StaleElementReferenceException occurred at line: {linha_erro}")
                                    if linha_erro >= 0 and linha_erro not in err[pagina - 1]:
                                        err[pagina-1].append(linha_erro)
                                        print(err[pagina - 1])
                                    self.driver.back()
                                    self.consulta_instrumento()
                                    self.loop_de_pesquisa(numero_processo=numero_processo,
                                                          caminho_pasta=caminho_pasta,
                                                          pasta_download=pasta_download,
                                                          feriado=feriado,
                                                          err=err,
                                                          pg=pagina,
                                                          anexos=anexos
                                                          )
                                except Exception as error:
                                    error_trace = traceback.format_exc()
                                    print(f'❌ Erro ao pular linha com falha. Erro:'
                                          f' {type(error).__name__}\nTraceback:\n{error_trace}')
                            except Exception as error:
                                print(f"❌ Erro ao processar a linha nº{indice} de termo, erro: {type(error).__name__}")
                                continue
                except Exception as error:
                    print(f"❌ Erro ao buscar nova página em anexos execução: {error}. Err:"
                          f" {type(error).__name__}")

            except Exception as error:
                print(f'❌ Erro de download: Termo {error}. Err: {type(error).__name__}')

        self.campo_pesquisa(numero_processo=numero_processo)

        try:
            # Executa pesquisa de anexos
            self.busca_propostas()

            # Seleciona lista de anexos execução e manda baixar os arquivos
            try:
                if not err:
                    err = [[]]
                    pg = 0
                print('\n🔁📎 Executando pesquisa de anexos execução')

                time.sleep(1)
                botao_lista_execucao = self.webdriver_element_wait('//tbody//tr//input[2]')

                if botao_lista_execucao.is_displayed() or botao_lista_execucao.is_enabled():
                    self.lista_execucao()
                    # Verifica se a tabela existe na página anexos execução.
                    try:
                        baixa_pdf_exec()
                    except Exception as e:
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

            except StaleElementReferenceException:
                try:
                    for _ in range(3):
                        self.driver.find_element(By.XPATH, '/html/body/div[3]/div[14]/div[3]/div[1]'
                                                   '/div/form/table/tbody/tr/td[2]/input[2]').click()
                        time.sleep(0.33)
                except Exception:
                    raise Exception("Stopping script execution due to an error.")
            except Exception as er:
                print(f'❌ Falha ao acessar documentos de execução.'
                      f'\nErro:{type(er).__name__}, {str(er)[:50]}')
                self.consulta_instrumento()

        except TimeoutException as t:
            print(f'TIMEOUT {t[:50]}')
            self.consulta_instrumento()
        except Exception as erro:
            print(f'❌ Falha ao acessar documentos de execução. Erro:{type(erro).__name__}\n{str(erro)[:50]}')
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
        def save_chrome_screen_as_pdf(max_attempts=3, print_delay=2, save_delay=3):
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
                    # Ensure Chrome is the active window
                    chrome_window = pyautogui.getWindowsWithTitle("Chrome")
                    if chrome_window:
                        chrome_window[0].activate()

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
                print(
                    f"❌ [File Search Error] Could not find or select PDF file: {type(e).__name__}: {str(e)[:50]}")
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
                print(
                    f"❌ [PDF Processing Error] Error during PDF cropping: {type(e).__name__}: {str(e)[:50]}")
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
            # Encontra a tabela de anexos
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

                    for indice in range(1, len(linhas)+1):
                        try:
                            # Refresh table and rows reference periodically
                            if indice > 0:
                                tabela = WebDriverWait(self.driver, 10).until(
                                    EC.presence_of_element_located((By.ID, 'esclarecimentos')))
                                linhas = tabela.find_elements(By.TAG_NAME, 'tr')
                        except Exception as error:
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

                            # Print page and handle subsequent actions
                            self.print_page(pasta_download=pasta_download, pdf_path=unique_pdf_path,)

                            # Use explicit wait for the final button
                            self.driver.execute_script(
                                "window.scrollTo(0, document.body.scrollHeight);")
                            pyautogui.hotkey('alt', 'left')

                        except Exception as error:
                            print(
                                f'❌ Erro ao imprimir a página. Erro: {type(error).__name__}\n'
                                f'Traceback:\n{str(error)[:50]}')
            except Exception as error:
                print(f"❌ Erro ao buscar nova página de esclarecimento: {error}.\nErr:"
                      f" {type(error).__name__}")
            self.consulta_instrumento()
        except Exception as error:
            print(f"❌ Erro ao econtrar elementos na página de esclarecimentos: {error}. Err:"
                  f" {type(error).__name__}")

        self.espera_completar_download(pasta_download=pasta_download)
        self.transfere_arquivos(caminho_pasta, pasta_download)
        print(
            f"\n{Fore.GREEN}✅ Loop de pesquisa concluído para o processo:"
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
                print(f"Falha com localizador {locator}: {str(e)[:80]}")
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

    # Pega a data do documento e o nome da aba
    def compara_data(self, data_site: str, feriado: int, no_comp: bool = True) -> bool:
        """
        Compara a data fornecida com a data de ontem.

        Args:
            data_site: Data a ser comparada, no formato 'AAAA-MM-DD' ou 'AAAA-MM-DD HH:MM:SS'.
            feriado: Quantidade de dias a serem descontadas para pegar os dias que o robo não rodou
            no_comp: Indica se usará comparação de data ou não

        Returns:
            True se a data fornecida for anterior à data de ontem, False caso contrário.
            Retorna False se o formato da data fornecida for inválido.
        """
        # Obtém a data de ontem
        try:
            if no_comp:
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
            exist_ok = True  # evita erro se a pasta já existir.
            os.makedirs(caminho_pasta, exist_ok=True)
            print(f"✅ Pasta '{nome_pasta}' criada em: {caminho_pasta}\n")
        except Exception as e:
            print(f"❌ Erro ao criar a pasta '{nome_pasta}': {e}")
        # Retorna o caminho completo da pasta, mesmo que a criação tenha falhado (para tratamento posterior)
        return caminho_pasta

    def extrair_dados_excel(self, caminho_arquivo_fonte, busca_id: str, tipo_instrumento_id: str):
        """
           Lê os contatos de uma planilha Excel e executa ações baseadas nos dados extraídos.

           Args:
               caminho_arquivo_fonte (str): Caminho do arquivo Excel que será lido.
               busca_id (str): Nome da coluna que contém os números de processo.
               tipo_instrumento_id (str): Nome da coluna que contém o tipo de processo.
           """
        pd.set_option('future.no_silent_downcasting', True)
        dados_processo = pd.read_excel(caminho_arquivo_fonte, dtype=str)
        dados_processo = dados_processo.replace({u'\xa0': ''})
        dados_processo = dados_processo.infer_objects(copy=False)

        # Cria um lista para cada coluna do arquivo xlsx
        numero_processo = list()
        tipo_instrumento = list()

        try:
            # Itera a planilha e armazena os dados em listas
            for indice, linha in dados_processo.iterrows():  # Assume que a primeira linha e um cabeçalho
                numero_processo.append(linha[busca_id])  # Busca o número do processo
                tipo_instrumento.append(linha[tipo_instrumento_id])  # Busca o tipo de instrumento
        except Exception as e:
            print(f"❌ Erro de leitura encontrado, erro: {e}")

        return numero_processo, tipo_instrumento

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


def main(feriado=2) -> None:
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
    caminho_arquivo_fonte = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e '
                             r'Assistência Social\Automações SNEAELIS\Analise_Custos_Exec_Print\source'
                             r'\RTMA Passivo 2024 - PROJETOS E PROGRAMAS 1.xlsx')
    # Rota da pasta onde os arquivos baixados serão alocados, cada processo terá uma subpasta dentro desta
    caminho_pasta_onedrive = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e '
                              r'Assistência Social\Automações SNEAELIS\Analise_Custos_Exec_Print')
    # Caminho do arquivo JSON que serve como catão de memória
    arquivo_log = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                   r'Social\Automações SNEAELIS\Analise_Custos_Exec_Print\source\arquivo_log.json')

    try:
        # Instancia um objeto da classe Robo
        robo = Robo()
        # Extrai dados de colunas específicas do Excel
    except Exception as e:
        print(f"\n‼️ Erro fatal ao iniciar o robô: {e}")
        sys.exit("Parando o programa.")

    numero_processo, tipo_instrumento = robo.extrair_dados_excel(
        caminho_arquivo_fonte=caminho_arquivo_fonte,
        busca_id='Instrumento nº',
        tipo_instrumento_id='Regime Jurídico do Instrumento (Modalidade)'
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
                feriado=feriado
            )
            try:
                robo.loop_esclarecimento(pasta_download=pasta_download,
                                         caminho_pasta=caminho_pasta,
                                         numero_processo=numero_processo[indice]
                                         )
            except Exception as e:
                print(f"\n❌ Erro ao processar esclarecimento no {indice}: ({numero_processo[indice]}).\n"
                      f" {e[:50]}")

            # Confirma se houve atualização na pasta e envia email para o técnico
            confirma_email = list(robo.condicao_email(numero_processo=numero_processo[indice],
                                                      caminho_pasta=caminho_pasta))
            if confirma_email:
                robo.salva_progresso(arquivo_log,numero_processo[indice], confirma_email[2], indice=indice)

        except Exception as e:
            print(f"\n❌ Erro ao processar o índice {indice} ({numero_processo[indice]}): {e}")
            robo.consulta_instrumento()
            # Você pode salvar o erro em log aqui, se quiser
            continue  # Continua para o próximo processo


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


def hash_file(file_path, block_size=65536):
    """Calculates SHA-256 hash of a file."""
    sha256 = hashlib.sha256()
    try:
        with open(file_path, 'rb') as f:
            for block in iter(lambda: f.read(block_size), b''):
                sha256.update(block)
        return sha256.hexdigest()
    except Exception as e:
        print(f"⚠️ Erro ao calcular hash de {file_path}: {e}")
        return None


def delete_duplicate_files(directory, recursive=False):
    """
    Deletes duplicate files in a directory based on file content (SHA-256 hash).

    Parameters:
        directory (str): Path to the target folder
        recursive (bool): Whether to search subdirectories
    """
    seen_hashes = {}
    deleted_files = []

    for dirpath, dirnames, filenames in os.walk(directory):
        for filename in filenames:
            full_path = os.path.join(dirpath, filename)
            file_hash = hash_file(full_path)

            if file_hash is None:
                continue  # Skip unreadable files

            if file_hash in seen_hashes:
                try:
                    os.remove(full_path)
                    deleted_files.append(full_path)
                    print(f"🗑️ Duplicado removido: {full_path}")
                except Exception as e:
                    print(f"❌ Erro ao deletar {full_path}: {e}")
            else:
                seen_hashes[file_hash] = full_path

        if not recursive:
            break  # Exit after the top-level directory

    print(f"\n✅ Concluído. {len(deleted_files)} arquivos duplicados deletados.")
    return deleted_files
if __name__ == "__main__":
    start_time = time.time()

    main()
    count_empty_folders(r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                        r'Social\Automações SNEAELIS\Analise_Custos_Exec_Print')

    end_time = time.time()
    tempo_total = end_time - start_time
    horas = int(tempo_total // 3600)
    minutos = int((tempo_total % 3600) // 60)
    segundos = int(tempo_total % 60)
    print(f'⏳ Tempo de execução: {horas}h {minutos}m {segundos}s')