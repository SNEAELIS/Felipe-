from selenium import webdriver
from selenium.common import WebDriverException, TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
#from datetime import datetime, timedelta
#from itertools import islice
import pandas as pd
import time
import os
import win32com.client as win32
import sys
#import shutil
#import json


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

    #
    def elemento_invisivel(self):
        try:
            WebDriverWait(self.driver, 30).until(
                EC.invisibility_of_element((By.CLASS_NAME, "backdrop")))
            print("🟢 Loading overlay disappeared.")
        except Exception as e:
            print(f"❌ Loading overlay did not disappear: {e}")

    # Consulta o instrumento ou proposta (faz todo o caminho até chegar no item desejado)
    def consulta_instrumento(self):
        """
               Navega pelas abas do sistema até a página de busca de processos.

               Esta função clica nas abas principal e secundária para acessar a página
               onde é possível realizar a busca de processos.
               """
        # Bloco de reset
        try:
            # Reseta para página inicial
            reset = self.driver.find_element(By.XPATH, "//a[normalize-space()='Transferências Especiais']")
            if reset:
                reset.click()
        except Exception as e:
            print(f'❌ Falha ao resetar, erro:')
            pass
        # Bloco de acesso ao processo
        try:
            # Seleciona o dropdown list
            self.webdriver_element_wait("//button[@title='Menu']").click()

            # Seleciona "PLANO DE AÇÃO"
            self.webdriver_element_wait("//span[normalize-space()='Plano de Ação']").click()
            print('🖱️✔️ Aba Plano de Ação selecionada ! \n')

            self.elemento_invisivel()

            # Abre a tabela de filtros
            self.webdriver_element_wait("//i[@class='fas fa-filter']").click()

            self.elemento_invisivel()

            # Seleciona o filtro de "Exibir planos de ação aguardando a análise do meu órgão"
            self.webdriver_element_wait("//label[@for='aguardandoUsuario']").click()

            # Aplica os filtros
            self.webdriver_element_wait("//button[@type='submit']").click()
            print('🖱️✔️ Filtro aplicado com sucesso ! \n')

            self.elemento_invisivel()

            # Extende a quantidade de itens por página
            exibir = self.webdriver_element_wait("/html/body/transferencia-especial-root/br"
                     "-main-layout/div/div/div/main/transferencia-especial-main/transferencia"
                     "-plano-acao/transferencia-plano-acao-consulta/br-table/div/ngx-"
                     "datatable/div/datatable-footer/div/br-pagination-table/div/br-select[1]/"
                     "div/div/ng-select/div/div/div[3]/input")
            exibir.click()
            exibir.send_keys('100', Keys.RETURN)

            self.elemento_invisivel()

            print(f"🖱️✔️ Sucesso em acessar a página de busca de processo ")

        except Exception as e:
            print(f'❌ Falha ao consultar o processo {e}')
            sys.exit(1)

    # Obtem o tamanho da tabela
    def obter_tamanho_tabela(self):
        """
        Obtém o tamanho (número de linhas) da primeira tabela.

        Retorna:
            tuple: Uma tupla contendo:
                - cabecalho (list): Lista dos cabeçalhos das colunas.
                - linhas (list): Lista das linhas da tabela.
        """
        try:
            # Localiza a tabela
            tabela = self.webdriver_element_wait("//div[@class='br-table is-grid-small']")

            # Extrai os cabeçalhos
            cabecalho = [cabeca.text for cabeca in tabela.find_elements(By.CSS_SELECTOR, "datatable-header-cell")]

            # Extrai as linhas
            linhas = tabela.find_elements(By.CSS_SELECTOR, "datatable-row-wrapper")
            print(f'📜📏 Quantidade de linhas encontradas: {len(linhas)}')
            return cabecalho, linhas
        except Exception as e:
            print(f"❌ Falha ao obter o tamanho da tabela: {e}")
            sys.exit(1)

    def iterar_tabela(self, numero_processo: list, caminho_arquivo: str):
        """
        Itera pela tabela para encontrar o processo desejado e extrair dados.

        Args:
            numero_processo (str): O número do processo a ser buscado.
            caminho_arquivo (str): O caminho para salvar o arquivo Excel.
        """
        dados_excel = []
        try:
            self.elemento_invisivel()
            # Obtém o tamanho e os dados da tabela
            cabecalho, linhas = self.obter_tamanho_tabela()

            # Itera pelas linhas
            for num, linha in enumerate(linhas):
                print(num+1)
                colunas = linha.find_elements(By.CSS_SELECTOR, "datatable-body-cell")
                dado_linha = {cabecalho[i]: colunas[i].text for i in range(len(cabecalho))}
                # Verifica se a linha atual corresponde ao processo desejado
                if dado_linha["Código"] in numero_processo:
                    valor_investimento = dado_linha['Valor']

                    # Abre os detalhes
                    linha.find_element(By.XPATH,
                           "/html[1]/body[1]/transferencia-especial-root[1]/br-main-layout[1]/div[1]"
                            "/div[1]/div[1]/main[1]/transferencia-especial-main[1]/transferencia-plano-acao[1]"
                            "/transferencia-plano-acao-consulta[1]/br-table[1]/div[1]/ngx-datatable[1]/div[1]"
                            "/datatable-body[1]/datatable-selection[1]/datatable-scroller[1]/datatable-row-"
                            f"wrapper[{num+1}]/datatable-body-row[1]/div[2]/datatable-body-cell[8]/div[1]/div[1]"
                            "/button[1]/i[1]"
                    ).click()

                    # Extrai o número do "PLANO DE AÇÃO"
                    plano_acao_numero = self.webdriver_element_wait(
                        "//div[@class='br-content']//div[1]//div[1]//div[1]//div[1]").text
                    beneficiario = self.webdriver_element_wait(
                        "//plano-acao-info//div[2]//div[1]//div[1]//div[1]").text
                    emenda_parlamentar = self.webdriver_element_wait(
                        "//plano-acao-info//div[2]//div[1]//div[2]//div[1]").text
                    print('Dados adquiridos com sucesso')

                    self.elemento_invisivel()

                    # Acessa a página "PLANO DE TRABALHO"
                    self.webdriver_element_wait("/html/body/transferencia-especial-"
                                     "root/br-main-layout/div/div/div/main/transferencia-especial-main"
                                     "/transferencia-plano-acao/transferencia-cadastro/br-tab-set/div/nav/ul/"
                                     "li[3]/button").click()
                    print('🖱️✔️ Aba Plano de Trabalho selecionada ! \n')

                    self.elemento_invisivel()

                    # Extrai dados necessários da aba Plano de Trabalho
                    situacao = self.driver.find_element(By.XPATH,
                        '/html/body/transferencia-especial-root/br-main-layout/div/div/div/main/'
                        'transferencia-especial-main/transferencia-plano-acao/transferencia-cadastro/br-tab'
                        '-set/div/nav/transferencia-plano-acao-plano-trabalho/form/div/div/br-fieldset[1]/'
                        'fieldset/div[2]/div[1]/div/br-input/div/div[1]/input').get_attribute("value")
                    objeto = ''
                    meta = ''
                    print(situacao, "situação ok")

                    # Accessa as metas
                    self.webdriver_element_wait(
                        "//i[@class='fas fa-chevron-up']").click()

                    # Extrai dados da tabela "LISTA DE EXECUTORES"
                    tabela_plano_trabalho = self.driver.find_element(By.CSS_SELECTOR, "br-table")
                    cabeca_plano_trabalho = [cab_plano_trabalho.text for cab_plano_trabalho in
                                            tabela_plano_trabalho.find_elements(By.CSS_SELECTOR,
                                            "datatable-header-cell")]
                    linhas_plano_trabalho = tabela_plano_trabalho.find_elements(By.CSS_SELECTOR,
                                                                      "datatable-row-wrapper")

                    for linha_plano_trabalho in linhas_plano_trabalho:
                        print(linhas_plano_trabalho)
                        colunas_exec = linha_plano_trabalho.find_elements(By.CSS_SELECTOR,
                                                                           "datatable-body-cell")
                        dado_linhas_plano_trabalho = {cabeca_plano_trabalho[i]: colunas_exec[i].text
                                                      for i in range(len(cabeca_plano_trabalho))}

                        if "Objeto" in dado_linhas_plano_trabalho:
                            objeto = dado_linhas_plano_trabalho["Objeto"]
                        elif "Executor" in dado_linhas_plano_trabalho:
                            if "Meta" in dado_linhas_plano_trabalho["Executor"]:
                                meta = dado_linhas_plano_trabalho


                    # Entra na aba Análises
                    analises = self.webdriver_element_wait("/html/body/transferencia-especial-root/br-main"
                               "-layout/div/div/div/main/transferencia-especial-main/transferencia-plano"
                               "-acao/transferencia-cadastro/br-tab-set/div/nav/ul/li[4]/button/span")
                    analises.click()
                    print('🖱️✔️ Aba Análises selecionada ! \n')

                    self.elemento_invisivel()

                    # Adiciona os dados extraídos à lista
                    dados_excel.append([plano_acao_numero,
                                        beneficiario,
                                        emenda_parlamentar,
                                        objeto,
                                        valor_investimento,
                                        situacao,

                                        ])

                    # Salva os dados no Excel
                    self.salva_excel(dados_excel, caminho_arquivo=caminho_arquivo)
                    self.webdriver_element_wait("//button[@id='btnVoltar']").click()
                else:
                    continue
        except NoSuchElementException:
            print('Element not found')
        except Exception as e:
            print(f'❌ Falha no loop de pesquisa do processo: {e}')
            sys.exit(1)

    # Salva os arquivos em uma planilha excel
    def salva_excel(self, dados: list, caminho_arquivo: str):
        nova_linha = pd.DataFrame(dados, columns=["Plano de Ação", "Objeto", "Valor Investimento"])

        # Busca a existência da planilha
        if os.path.exists(caminho_arquivo):
            # Carrega a planilha se já existir
            df_existente = pd.read_excel(caminho_arquivo)
            df_combinado = pd.concat([df_existente, nova_linha], ignore_index=True)
        else:
            df_combinado = nova_linha

        # Salva o arquivo em xlsx
        df_combinado.to_excel(caminho_arquivo, index=False)
        print(f"✅ Dados salvos/atualizados em {os.path.basename(caminho_arquivo)}")

    def passar_pagina(self):
        self.webdriver_element_wait("//button[@id='btn-next-page']").click()

    # Extrai dados da planílha de controle. Planilha gerada com os filtros todos ativados.
    def extrair_dados_excel(self, caminho_arquivo_fonte, busca_id: str):
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

        try:
            # Itera a planilha e armazena os dados em listas
            for indice, linha in dados_processo.iterrows():  # Assume que a primeira linha e um cabeçalho
                numero_processo.append(linha[busca_id])  # Busca o número do processo

        except Exception as e:
            print(f"❌ Erro de leitura encontrado, erro: {e}")

        return numero_processo

def main() -> None:
    # Caminho do arquivo .xlsx que contem os dados necessários para rodar o robô
    caminho_arquivo_fonte = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e'
                              r' Assistência Social\Automações '
                             r'SNEAELIS\Robo Transferências Especiais\filtro.xlsx')

    # Rota da pasta onde os arquivos baixados serão alocados, cada processo terá uma subpasta dentro desta
    caminho_arquivo = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e'
                       r' Assistência Social\Automações SNEAELIS\Robo Transferências Especiais'
                       r'\resultado_robo.xlsx')

    # Instancia um objeto da classe Robo
    robo = Robo()

    # Extrai dados de colunas específicas do Excel
    numero_processo = robo.extrair_dados_excel(caminho_arquivo_fonte, 'Código')

    robo.consulta_instrumento()
    while True:
        try:
            robo.iterar_tabela(caminho_arquivo=caminho_arquivo, numero_processo=numero_processo)
            robo.passar_pagina()
        except Exception:
            break

if __name__ == "__main__":
    start_time = time.time()
    main()
    end_time = time.time()
    tempo_total = end_time - start_time
    horas = int(tempo_total // 3600)
    minutos = int((tempo_total % 3600) // 60)
    segundos = int(tempo_total % 60)
    print(f'⏳ Tempo de execução: {horas}h {minutos}m {segundos}s')

