import time
import os
import sys
import json
import shutil

import win32com.client as win32

import pandas as pd

from selenium import webdriver
from selenium.common import WebDriverException, TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver import ActionChains

from webdriver_manager.chrome import ChromeDriverManager

from datetime import datetime, timedelta

from colorama import init, Fore, Back, Style


class ProgressMonitor:
    def __init__(self, progress_file="scraping_progress.json"):
        self.progress_file = progress_file
        self.last_index = -1
        self.last_process = None
        self.completed_processes = []

    def load_progress(self):
        """Load saved progress from JSON file"""
        try:
            if os.path.exists(self.progress_file):
                with open(self.progress_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.last_index = data.get('last_index', -1)
                    self.last_process = data.get('last_process', None)
                    self.completed_processes = data.get('completed_processes', [])
                    print(f"📂 Progresso carregado: Último índice = {self.last_index}, Processo = {self.last_process}")
                    return True
            else:
                print("📂 Nenhum progresso anterior encontrado. Iniciando do início.")
                return False
        except Exception as e:
            print(f"⚠️ Erro ao carregar progresso: {e}")
            return False

    def save_progress(self, index, process_number):
        """Save current progress to JSON file"""
        try:
            self.last_index = index
            self.last_process = process_number

            # Add to completed processes if not already there
            if process_number not in self.completed_processes:
                self.completed_processes.append(process_number)

            data = {
                'last_index': self.last_index,
                'last_process': self.last_process,
                'completed_processes': self.completed_processes,
                'total_processed': len(self.completed_processes),
                'last_update': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }

            with open(self.progress_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=4, ensure_ascii=False)

            print(f"💾 Progresso salvo: Índice {index} - Processo {process_number}")
            return True
        except Exception as e:
            print(f"⚠️ Erro ao salvar progresso: {e}")
            return False

    def reset_progress(self, source_path, destiny_path):
        """Reset all progress"""
        try:
            shutil.copy2(src=source_path, dst=destiny_path)
            self.last_index = -1
            self.last_process = None
            self.completed_processes = []

            if os.path.exists(self.progress_file):
                os.remove(self.progress_file)

            print("🔄 Progresso resetado com sucesso!")
            return True
        except Exception as e:
            print(f"⚠️ Erro ao resetar progresso: {e}")
            return False

    def get_start_index(self):
        """Get the index to start from (last_index + 1)"""
        return self.last_index + 1

    def show_summary(self, total_items):
        """Show progress summary"""
        print("\n" + "=" * 70)
        print("📊 RESUMO DO PROGRESSO".center(70))
        print("=" * 70)
        print(f"📌 Total de processos: {total_items}")
        print(f"✅ Processos concluídos: {len(self.completed_processes)}")
        print(f"📈 Progresso: {len(self.completed_processes) / total_items * 100:.1f}%")
        print(f"📍 Último índice processado: {self.last_index}")
        print(f"🔖 Último processo: {self.last_process}")
        print(f"🕒 Última atualização: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        print("=" * 70)


def ask_restart_option():
    """Ask user how to proceed with the scraping"""
    print("\n" + "=" * 70)
    print("🔄 OPÇÕES DE EXECUÇÃO".center(70))
    print("=" * 70)
    print("1. 🔄 Continuar do último índice salvo")
    print("2. 🆕 Resetar progresso e começar do início")
    print("3. 📊 Apenas mostrar resumo e sair")
    print("4. 🎯 Iniciar de um índice específico")
    print("=" * 70)

    while True:
        try:
            choice = input("\n👉 Escolha uma opção (1-4): ").strip()

            if choice == '1':
                print("\n✅ Continuando do último progresso salvo...")
                return 'continue'
            elif choice == '2':
                print("\n🔄 Resetando progresso e iniciando do início...")
                return 'reset'
            elif choice == '3':
                print("\n📊 Exibindo apenas o resumo...")
                return 'summary'
            elif choice == '4':
                return 'specific'
            else:
                print("⚠️ Opção inválida. Escolha 1, 2, 3 ou 4.")
        except KeyboardInterrupt:
            print("\n\n⚠️ Execução cancelada pelo usuário.")
            sys.exit(0)
        except Exception as e:
            print(f"❌ Erro: {e}")


def ask_specific_index(max_index):
    """Ask user for a specific starting index"""
    while True:
        try:
            index_input = input(f"\n🎯 Digite o índice inicial (0 a {max_index - 1}): ").strip()
            specific_index = int(index_input)

            if 0 <= specific_index < max_index:
                print(f"✅ Iniciando do índice {specific_index}")
                return specific_index
            else:
                print(f"⚠️ Índice inválido. Digite um número entre 0 e {max_index - 1}")
        except ValueError:
            print("⚠️ Por favor, digite um número válido.")
        except KeyboardInterrupt:
            print("\n\n⚠️ Execução cancelada pelo usuário.")
            sys.exit(0)



class Robo:
    def __init__(self):
        """
        Inicializa o objeto Robo, configurando e iniciando o driver do Chrome.
        """
        try:
            try:
                # Inicia as opções do Chrome
                self.chrome_options = webdriver.ChromeOptions()
                # Endereço de depuração para conexão com o Chrome
                self.chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
                # Inicializa o driver do Chrome com as opções e o gerenciador de drivers
                self.driver = webdriver.Chrome(
                    service=Service(ChromeDriverManager().install()),
                    options=self.chrome_options)
            except Exception as e:
                print(f"Error with ChromeDriverManager: {type(e).__name__}.\n Erro: {str(e)[:80]}")
                sys.exit()

            qnt_abas = self.skip_chrome_tab_search()

            self.driver.switch_to.window(qnt_abas[0])


            print(f"✅ Controle sobre a na página {self.driver.title}.")

        except WebDriverException as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            # Imprime mensagem de erro se a conexão falhar
            print(f"❌ Erro ao conectar ao navegador existente: {type(e).__name__}.\n Erro: {str(e)[:80]}")
            # Define o driver como None em caso de falha na conexão
            self.driver = None

    
    def skip_chrome_tab_search(self):
        qnt_abas = self.driver.window_handles
        abas_limpas = qnt_abas
        for handle in qnt_abas:
            self.driver.switch_to.window(handle)
            url = self.driver.current_url
            if "chrome" in url:
                abas_limpas.remove(handle)

        return  abas_limpas


    # Chama a função do webdriver com wait element to be clickable
    def webdriver_element_wait(self, xpath: str):
        # Cria uma instância de WebDriverWait com o driver e o tempo limite e espera o elemento ser clicável
        return WebDriverWait(self.driver, 20).until(EC.element_to_be_clickable((By.XPATH, xpath)))


    # Navega até a página de busca do instrumento
    def consulta_instrumento(self):
        # Reseta para página inicial
        try:
            reset = self.webdriver_element_wait('//*[@id="header"]')
            if reset:
                img = reset.find_element(By.XPATH, '//*[@id="logo"]/a/img')
                action = ActionChains(self.driver)
                action.move_to_element(img).perform()
                reset.find_element(By.TAG_NAME, 'a').click()
                print(Fore.MAGENTA + "\n✅ Processo resetado com sucesso !")
        except:
            pass

        # [0] Execução; [1] Consultar Instrumentos/Pré-Instrumentos
        xpaths = ['//*[@id="menuPrincipal"]/div[1]/div[4]',
                  '//*[@id="contentMenu"]/div[1]/ul/li[6]/a'
                  ]
        try:
            for idx in range(len(xpaths)):
                self.webdriver_element_wait(xpaths[idx]).click()
            print(f"{Fore.MAGENTA}✅ Sucesso em acessar a página de busca de processo{Style.RESET_ALL}")

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(Fore.RED + f'🔴📄 Instrumento indisponível. \nErro: {type(e).__name__}.\n Erro: {str(e)[:80]}')
            sys.exit(1)


    def campo_pesquisa(self, numero_processo):
        try:
            # Seleciona campo de consulta/pesquisa, insere o número de proposta/instrumento e da ENTER
            campo_pesquisa = self.webdriver_element_wait('//*[@id="consultarNumeroConvenio"]')
            campo_pesquisa.clear()
            campo_pesquisa.send_keys(numero_processo)
            campo_pesquisa.send_keys(Keys.ENTER)

            # Acessa o item proposta/instrumento
            acessa_item = self.webdriver_element_wait('//*[@id="instrumentoId"]/a')
            acessa_item.click()
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f' Falha ao inserir número de processo no campo de pesquisa. Erro: {type(e).__name__}')


    # Confere qual o status mais atual do instrumento
    def check_situacao(self, path) -> bool:
        try:
            self.webdriver_element_wait(path)
            tabela = self.driver.find_element(By.XPATH, path)
            linhas = tabela.find_elements(By.CSS_SELECTOR, '#tbodyrow tr')

            if linhas:
                ultima_linha = linhas[-1]

                colunas = ultima_linha.find_elements(By.TAG_NAME, 'td')

                situacao_texto = colunas[1].text.strip().lower()
                print(situacao_texto)
                if 'Em Análise'.lower() in situacao_texto or 'Cadastrado'.lower() in situacao_texto:
                    colunas[3].find_element(By.LINK_TEXT, "Detalhar").click()
                    return True

            return False
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f"⚠️ Erro ao verificar situação: {type(e).__name__}")
            return False


    def extrair_dado_detalhado(self, aba: str) -> dict | None:
        try:
            page_path = '//*[@id="editar"]/div/form/table'
            tr = self.webdriver_element_wait(page_path)

            labels = [el.text for el in tr.find_elements(By.CSS_SELECTOR, "td.label")]
            fields = [el.text for el in tr.find_elements(By.CSS_SELECTOR, "td.field")]

            raw_data = {str(label).strip(): str(field).strip() for label, field in zip(labels, fields) if label.strip()}
            print({f'Data da Solicitação {aba}': raw_data['Data da Solicitação']})
            btn_voltar = self.driver.find_elements(By.XPATH, '//*[@id="form_submit"]')
            try:
                for _ in btn_voltar:
                    if _.get_attribute('value').strip() == 'Voltar':
                        btn_voltar.click()
                        return {f'Data da Solicitação {aba}': raw_data['Data da Solicitação']}
            except:
                pass

            if raw_data:
                return {f'Data da Solicitação {aba}': raw_data['Data da Solicitação']}
            else:
                return {f'Data da Solicitação {aba}': ''}
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Erro ao extrair dado detalhado para {aba}: {type(e).__name__}")
            return {}


    def dados_pt(self) -> dict | None:
        print(f'\n📤📄 Extraindo dados da pagina "AJUSTES DO PT"')

        xpaths = ['//*[@id="menu_link_-481524888_-1293190284"]',  # Aba "AJUSTES DO PT"
                  '//*[@id="listaAjustePlanoTrabalho"]',  # Tabela "AJUSTES DO PT"
                  ]
        try:
            self.webdriver_element_wait(xpaths[0]).click()

            sitacao = self.check_situacao(xpaths[1])
            dados_pt = {'Data da Solicitação PT':''}
            if sitacao:
                dados_pt = self.extrair_dado_detalhado('PT')
                return dados_pt
            else:
                return dados_pt
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Erro ao extrair dados do PT: {type(e).__name__}")


    def tas(self) -> dict | None:
        print(f'📤📄 Extraindo dados da pagina "TAs"')

        xpaths = ['//*[@id="menu_link_-481524888_82854"]',  # Aba "TAs"
                  '//*[@id="listaSolicitacoes"]',  # Tabela "TAs"
                  ]
        try:
            self.webdriver_element_wait(xpaths[0]).click()

            sitacao = self.check_situacao(xpaths[1])
            dados_ta = {'Data da Solicitação TA':''}
            if sitacao:
                dados_ta = self.extrair_dado_detalhado('TA')
                return dados_ta
            else:
                return dados_ta
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Erro ao extrair dados do TA: {type(e).__name__}")


    def rendimento(self) -> dict | None:
        print(f'📤📄 Extraindo dados da pagina "RENDIMENTO DE APLICAÇÃO"')

        xpaths = ['//*[@id="menu_link_-481524888_1776368057"]',  # Aba "RENDIMENTO DE APLICAÇÃO"
                  '//*[@id="listaSolicitacoes"]',  # Tabela "RENDIMENTO DE APLICAÇÃO"
                  ]
        try:
            self.webdriver_element_wait(xpaths[0]).click()

            sitacao = self.check_situacao(xpaths[1])
            dados_rend = {'Data da Solicitação Rendimento':''}
            if sitacao:
                dados_rend = self.extrair_dado_detalhado('Rendimento')
                return dados_rend
            else:
                return dados_rend
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Erro ao extrair dados de rendimento: {type(e).__name__}")


    def anexos_exec(self) -> dict | None:
        print(f'\n📤📄 Extraindo dados da pagina "ANEXOS EXECUÇÃO"')

        try:
            xpaths = ['//*[@id="menu_link_997366806_1965609892"]',  # Aba "ANEXOS"
                      '//*[@id="form_submit"]'  # Botão "LISTAR ANEXOS EXECUÇÃO"
                      ]
            self.webdriver_element_wait(xpaths[0]).click()
            self.webdriver_element_wait(xpaths[1])
            lista_exec = self.driver.find_elements(By.XPATH, xpaths[1])
            print([t.get_attribute('value') for t in lista_exec])
            try:
                for _ in lista_exec:
                    if _.get_attribute('value') == 'Listar Anexos Execução':
                        _.click()
            except StaleElementReferenceException:
                try:
                    for _ in range(3):
                        self.webdriver_element_wait(xpaths[1])
                        lista_exec = self.driver.find_elements(By.XPATH, xpaths[1])
                        for _ in lista_exec:
                            if _.get_attribute('value') == 'Listar Anexos Execução':
                                _.click()
                except:
                    pass
            except:
                pass

            tabela = self.webdriver_element_wait('//*[@id="tituloListagem"]')
            linhas = tabela.find_elements(By.CSS_SELECTOR, '#tbodyrow tr')

            primeira_linha = linhas[0]
            colunas = primeira_linha.find_elements(By.TAG_NAME, 'td')
            data_anexo_exec = colunas[2].text.strip()

            return {'Anexos Execução': data_anexo_exec}
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Erro ao extrair anexos de execução: {type(e).__name__}")
            return {'Anexos Execução': ''}


    def salvar_dados_extracao(self, lista_dados, instrumento_numero, caminho_excel):
        try:
            # Initialize the row dictionary with the instrument number
            nova_linha = {'Instrumento Nº': instrumento_numero}

            # Extract data from each dictionary in the list
            for item in lista_dados:
                if isinstance(item, dict):
                    for key, value in item.items():
                        nova_linha[key] = value if pd.notna(value) and value != '' else None

            # Check if the Excel file already exists
            if os.path.exists(caminho_excel):
                # Read existing data
                df_existente = pd.read_excel(caminho_excel, dtype=str)

                # Check if instrument already exists
                if 'Instrumento Nº' in df_existente.columns:
                    if instrumento_numero in df_existente['Instrumento Nº'].values:
                        # Update existing row
                        idx = df_existente[df_existente['Instrumento Nº'] == instrumento_numero].index[0]
                        for key, value in nova_linha.items():
                            if key != 'Instrumento Nº':
                                df_existente.at[idx, key] = value
                        df_novo = df_existente
                        print(f"✅ Instrumento {instrumento_numero} atualizado no arquivo existente")
                    else:
                        # Append new row
                        df_novo = pd.concat([df_existente, pd.DataFrame([nova_linha])], ignore_index=True)
                        print(f"✅ Novo instrumento {instrumento_numero} adicionado ao arquivo existente")
                else:
                    # If 'Instrumento Nº' column doesn't exist, just append
                    df_novo = pd.concat([df_existente, pd.DataFrame([nova_linha])], ignore_index=True)
                    print(f"✅ Dados adicionados ao arquivo existente")
            else:
                # Create new DataFrame
                df_novo = pd.DataFrame([nova_linha])
                # Ensure directory exists
                os.makedirs(os.path.dirname(caminho_excel), exist_ok=True)
                print(f"✅ Novo arquivo criado em: {caminho_excel}")

            # Save to Excel
            df_novo.to_excel(caminho_excel, index=False, engine='openpyxl')
            print(f"📊 Dados salvos com sucesso!")

            return True

        except Exception as e:
            print(f"❌ Erro ao salvar dados no Excel: {e}")
            return False


    # Pesquisa o termo de fomento listado na planilha e executa download e transferência caso exista algúm.
    def loop_de_pesquisa(self, numero_processo: str, caminho_excel: str):
        todos_dados = []
        self.campo_pesquisa(numero_processo)
        try:
            self.webdriver_element_wait('//*[@id="tabControls"]') # Aba "EXECUÇÃO CONVENENTE"
            for _ in self.driver.find_elements(By.TAG_NAME, 'a'):
                if _.text.strip().lower() == 'Execução Convenente'.strip().lower():
                    _.click()
                    break
            else:
                sys.exit()
        except Exception as error:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f'❌ Err: {type(error).__name__}\n{str(error)[:80]}. ')

        try:
            todos_dados.append(self.dados_pt())
            todos_dados.append(self.tas())
            todos_dados.append(self.rendimento())

            xpath_plano_trab = '//*[@id="grupo_abas"]/a[2]'  # Aba "PLANO DE TRABALHO"
            self.webdriver_element_wait(xpath_plano_trab).click()

            todos_dados.append(self.anexos_exec())

            self.driver.find_element(By.XPATH, '//*[@id="breadcrumbs"]/a[2]').click()

            self.salvar_dados_extracao(lista_dados=todos_dados, instrumento_numero=numero_processo, caminho_excel=caminho_excel)

        except Exception as error:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f'❌ Erro de download. Err: {type(error).__name__}\n{str(error)[:80]}. ')


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
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            paginas = 1
            return paginas
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f"⚠️ Erro ao contar páginas: {type(e).__name__}")
            return 1


    # Pega a data do documento e o nome da aba
    def compara_data(self, data_site: str, feriado: int) -> bool:
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
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Erro na comparação de datas")
            return False


    def converter_data(self, data: str) -> datetime | None:
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


    def extrair_dados_excel(self, caminho_arquivo_fonte, busca_id: str):
        try:
            dados_processo = pd.read_excel(caminho_arquivo_fonte, dtype=str)
            dados_processo = dados_processo.replace(u'\xa0', '', regex=True)
            dados_processo = dados_processo.infer_objects(copy=False)

            dados_processo = self.filter_by_column(df=dados_processo,
                                  column='Status',
                                  condition='eq',
                                  value='ATIVOS TODOS'
                                  )

            # Cria um lista para cada coluna do arquivo xlsx
            numero_processo = list()

            try:
                # Itera a planilha e armazena os dados em listas
                for indice, linha in dados_processo.iterrows():  # Assume que a primeira linha e um cabeçalho
                    numero_processo.append(linha[busca_id])  # Busca o número do processo
            except Exception as e:
                exc_type, exc_value, exc_tb = sys.exc_info()
                print(f"\nError occurred at line: {exc_tb.tb_lineno}")
                print(f"❌ Erro de leitura encontrado, erro: {type(e).__name__}.\n Erro: {str(e)[:80]}")

            return numero_processo
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Erro ao extrair dados do Excel: {type(e).__name__}")
            return []


    # Limpa os dados que vem da planilha
    def limpa_dados(self, lista: list):
        try:
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
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Erro ao limpar dados: {type(e).__name__}")
            return lista


    # Verifica as condições para mandar um e-mail para o técnico
    def condicao_email(self, caminho_pasta: str):
        """Compara dois arquivos .xlsx na pasta, retorna um dicionário com as mudanças de data
        e os cabeçalhos (última palavra) das colunas preenchidas nas linhas alteradas."""
        try:
            # 1. Localiza os dois arquivos .xlsx mais recentes
            xlsx_files = [
                f for f in os.listdir(caminho_pasta)
                if f.endswith('.xlsx') and os.path.isfile(os.path.join(caminho_pasta, f))
            ]
            if len(xlsx_files) < 2:
                print(f"⚠️ Menos de 2 arquivos Excel encontrados.\n")
                return {}

            # Ordena por data de modificação (mais antigo primeiro)
            xlsx_files.sort(key=lambda f: os.path.getmtime(os.path.join(caminho_pasta, f)))
            arquivo_antigo = os.path.join(caminho_pasta, xlsx_files[-2])
            arquivo_novo = os.path.join(caminho_pasta, xlsx_files[-1])

            # 2. Leitura tratando tudo como string inicialmente
            df_old = pd.read_excel(arquivo_antigo, dtype=str)
            df_new = pd.read_excel(arquivo_novo, dtype=str)

            if df_old.empty or df_new.empty:
                print(f"⚠️ Um dos arquivos está vazio.\n")
                return {}

            # Coluna A = identificador (primeira coluna), Coluna B = data (segunda coluna)
            col_id = df_old.columns[0]
            col_data = df_old.columns[1]

            # Converte colunas de data para datetime (para comparação confiável)
            df_old[col_data] = pd.to_datetime(df_old[col_data], errors='coerce')
            df_new[col_data] = pd.to_datetime(df_new[col_data], errors='coerce')

            # 3. Merge pela chave e identifica diferenças
            merged = df_old.merge(df_new, on=col_id, suffixes=('_old', '_new'), how='inner')
            data_old = col_data + '_old'
            data_new = col_data + '_new'

            # Máscara considerando NaT como diferente de qualquer data
            mask_diff = (
                    merged[data_old].fillna(pd.Timestamp('1900-01-01'))
                    != merged[data_new].fillna(pd.Timestamp('1900-01-01'))
            )
            mudancas = merged[mask_diff]

            if mudancas.empty:
                print(f"⚠️ Nenhuma mudança de data encontrada.\n")
                return {}

            # 4. Monta o dicionário de resultado
            resultado = {}
            for _, row in mudancas.iterrows():
                chave = row[col_id]
                data_antiga = row[data_old]
                data_nova = row[data_new]

                # Linha correspondente no arquivo NOVO
                linha_nova = df_new[df_new[col_id] == chave].iloc[0]

                colunas_preenchidas = []
                for col in df_new.columns:
                    if col == col_id or col == col_data:
                        continue
                    valor_celula = linha_nova[col]
                    # Considera preenchido se não for nulo e não for string vazia
                    if pd.notna(valor_celula) and str(valor_celula).strip() != '':
                        ultima_palavra = col.split()[-1] if col.split() else col
                        colunas_preenchidas.append(ultima_palavra)

                resultado[chave] = {
                    'old_date': str(data_antiga.date()) if pd.notna(data_antiga) else 'N/A',
                    'new_date': str(data_nova.date()) if pd.notna(data_nova) else 'N/A',
                    'filled_columns': colunas_preenchidas
                }

            print(f"📂✨ Mudanças encontradas")
            return resultado

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nErro na linha {exc_tb.tb_lineno}: {exc_value}")
            print(f"⚠️ Falha ao comparar arquivos. {type(e).__name__}\n{str(e)[:100]}")
            return {}


    @staticmethod
    def filter_by_column(df, column, condition, value=None):
        """
        Filter a DataFrame based on a condition applied to a single column.

        Parameters:
        -----------
        df : pandas.DataFrame
            The DataFrame to filter
        column : str
            The column name to apply the condition on
        condition : str
            The condition type. Options: 'eq', 'ne', 'gt', 'lt', 'ge', 'le',
            'contains', 'startswith', 'endswith', 'in', 'not_in', 'isna', 'not_na'
        value : any, optional
            The value to compare against. Required for most conditions except 'isna' and 'not_na'

        Returns:
        --------
        pandas.DataFrame
            Filtered DataFrame
        """
        try:
            if column not in df.columns:
                raise ValueError(f"Column '{column}' not found in DataFrame")

            conditions = {
                'eq': df[column] == value,
                'ne': df[column] != value,
                'gt': df[column] > value,
                'lt': df[column] < value,
                'ge': df[column] >= value,
                'le': df[column] <= value,
                'contains': df[column].astype(str).str.contains(value, na=False),
                'startswith': df[column].astype(str).str.startswith(value, na=False),
                'endswith': df[column].astype(str).str.endswith(value, na=False),
                'in': df[column].isin([value]),
                'not_in': ~df[column].isin([value]),
                'isna': df[column].isna(),
                'not_na': df[column].notna()
            }

            if condition not in conditions:
                raise ValueError(f"Condition '{condition}' not recognized. "
                                 f"Available conditions: {list(conditions.keys())}")

            # For conditions that don't require value
            if condition in ['isna', 'not_na'] and value is not None:
                print(f"Warning: 'value' parameter ignored for condition '{condition}'")

            # For conditions that require value
            if condition not in ['isna', 'not_na'] and value is None:
                raise ValueError(f"Condition '{condition}' requires a 'value' parameter")

            mask = conditions[condition]
            return df[mask].copy()
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Erro ao filtrar dados do Excel: {type(e).__name__}\n{str(e)[:100]}")
            sys.exit()


def send_emails_from_excel():
    """
    Prepares emails from Excel data (Columns A, B, C) and opens them in Outlook for review.

    Args:
        excel_path (str): Path to Excel file.
        attachment_path (str, optional): Attachment file path.
        sender (str, optional): Email sender address.
    """

    def prepare_outlook_email(email, subject, html_body):
        """Creates and displays an Outlook email (without sending)."""
        try:
            if not all(isinstance(x, str) for x in [email, subject, html_body]):
                raise TypeError("Recipient, subject, and body must be strings")

            outlook = win32.Dispatch('outlook.application')
            time.sleep(2)
            e_mail = outlook.CreateItem(0)

            e_mail.To = email
            e_mail.Subject = subject
            e_mail.HTMLBody = html_body

            e_mail.Display()  # Opens email for review (instead of .Send())
            print(f"📧 Email prepared for: {email}")
            return True
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Failed to prepare email for {email}: \n{e}")
            return False

    def generate_email_body(extra_data_list):
        """Generates HTML body from Excel data only."""
        today = datetime.now().strftime("%d/%m/%Y")
        subject = f"Atualização do Acompanhamento de instrumentos. Data {today}"
        html_body = f"""<!DOCTYPE html>
        <html>
        <head>
            <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        </head>
        <body style="font-family: 'Courier New', monospace; margin: 0; padding: 0;">
        <div style="padding: 10px; background-color: #f0f0f0; border-bottom: 1px solid #ccc;
         font-weight: bold;">
        </div>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
                <td>"""
        for data in extra_data_list:
           ''' html_body += (
                f"<p style='white-space: pre ; font-family:Courier New, monospace>Código do Plano de "
                f"Ação: {data['A']}      "
                f"Responsável:"
                f" {data['B']}    "
                f"Data/Hora: {data['C']}    Situação: {data['D']}</p>"
                "<br><br></p>"  # 2 empty lines between entries
            )'''
           html_body += f"""
                       <div style="white-space: pre; margin-bottom: 20px;">
           Código do Plano de Ação:    {data['A']}
           Responsável:               {data['B']}
           Data/Hora:                 {data['C']}
           Situação:                  {data['D']}
                       </div>
                       """

           # Close all tags
           html_body += """</td>
               </tr>
            </table>
           </body>
           </html>"""
        return html_body, subject

    def read_excel_data():
        """Reads Excel and extracts Columns A, B, C if C has data."""
        try:
            email = ""
            extra_data_list = []  # Stores {A, B, C} dicts

            for _, row in df.iterrows():
                # Check if Column C (index 2) has data
                if pd.notna(row.iloc[3]):
                    extra_data = {
                        'A': row.iloc[0],  # Column A
                        'B': row.iloc[1],  # Column B
                        'C': row.iloc[2],  # Column C
                        'D': row.iloc[3]  # Column D
                    }
                    extra_data_list.append(extra_data)

            print(f'Número de planos de ação no email: {len(extra_data_list)}')
            if extra_data_list:
                return email, extra_data_list
            else:
                return [], []

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Failed to read Excel file: \n{e}")
            return [], []

    email_paths = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\Acompanhamento\E-mails_Acompanhamento.xlsx"

    excel_path_last_result = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Acompanhamento\Result_Scrape_Acomp_last_result.xlsx"

    excel_path_current_result = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Acompanhamento\Result_Scrape_Acomp.xlsx"

    # Read data
    emails, previous_data, current_data = read_excel_data()

    try:
        if email or extra_data_list:
        # Prepare emails
            html_body, subject = generate_email_body(extra_data_list)  # Pass single entry
            prepare_outlook_email(email, subject, html_body)
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"Error occurred at line: {exc_tb.tb_lineno}")
        print(f"❌ No data to prepare email for {type(e).__name__}: \n{str(e)[:100]}")


def main() -> None:
    def eta(indice, max_linha, start_time, bar_length=20):
        try:
            idx = indice + 1
            elapsed = time.time() - start_time

            if idx == 0:
                return

            # Average time per iteration
            avg = elapsed / idx

            # Remaining time
            remaining = max(max_linha - idx, 0)
            eta_sec = remaining * avg

            # Format time (HH:MM:SS)
            eta_h = int(eta_sec // 3600)
            eta_m = int((eta_sec % 3600) // 60)
            eta_s = int(eta_sec % 60)

            # Progress %
            progress = idx / max_linha if max_linha else 0

            # Progress bar
            filled = int(bar_length * progress)
            bar = "█" * filled + "-" * (bar_length - filled)

            print(
                f"[{bar}] {progress * 100:6.2f}% "
                f"| ETA {eta_h:02d}:{eta_m:02d}:{eta_s:02d} "
                f"| {idx}/{max_linha}",
                end="",
                flush=True
            )
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f"⚠️ Erro no cálculo do ETA: {type(e).__name__}")

    # Caminho do arquivo .xlsx que contem os dados necessários para rodar o robô
    caminho_arquivo_fonte = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\Acompanhamento\CONTROLE_DE_PARCERIAS_CGAP_28-04.xlsx"

    # Caminho do arquivo .xlsx que contem os dados da extração atual do robô
    caminho_arquivo_destino = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Acompanhamento\Results\Result_Scrape_Acomp.xlsx"

    # Caminho do arquivo .xlsx que contem cópia dos dados da ultima extração do robô
    caminho_arquivo_cópia = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Acompanhamento\Results\Result_Scrape_Acomp_last_result.xlsx"


    # Initialize progress monitor
    monitor = ProgressMonitor("scraping_progress_analise_custos.json")

    # Load previous progress
    monitor.load_progress()

    # Ask user how to proceed
    restart_option = ask_restart_option()

    try:
        robo = Robo()
    except Exception as e:
        print(f"\n‼️ Erro fatal ao iniciar o robô: {type(e).__name__}\n{str(e)[:100]}")
        sys.exit("Parando o programa.")

    try:
        numero_processo = robo.extrair_dados_excel(
            caminho_arquivo_fonte=caminho_arquivo_fonte,
            busca_id='Instrumento nº',
        )
    except Exception as e:
        print(f"❌ Erro ao extrair dados do Excel: {e}")
        sys.exit("Parando o programa.")

    # Pós-processamento dos dados

    numero_processo = robo.limpa_dados(numero_processo)
    max_linha = len(numero_processo)

    if restart_option == 'summary':
        monitor.show_summary(max_linha)
        print("\n👋 Encerrando programa...")
        sys.exit(0)
    elif restart_option == 'reset':
        monitor.reset_progress(source_path=caminho_arquivo_destino, destiny_path=caminho_arquivo_cópia)
        start_index = 0
    elif restart_option == 'specific':
        start_index = ask_specific_index(max_linha)
        # Optionally update the monitor to start from this index
        monitor.last_index = start_index - 1
        monitor.save_progress(monitor.last_index, "manual_start")
    else:  # 'continue'
        start_index = monitor.get_start_index()

        # Check if all processes are already done
        if start_index >= max_linha:
            print("\n✅ Todos os processos já foram concluídos!")
            monitor.show_summary(max_linha)
            sys.exit(0)

    # Show starting summary
    print("\n" + "=" * 70)
    print("🚀 INICIANDO EXECUÇÃO".center(70))
    print("=" * 70)
    print(f"📌 Total de processos: {max_linha}")
    print(f"📍 Iniciando do índice: {start_index}")
    print(f"📝 Processo inicial: {numero_processo[start_index] if start_index < max_linha else 'N/A'}")
    print(f"📊 Processos restantes: {max_linha - start_index}")
    print("=" * 70 + "\n")

    # Start the scraping
    start_time = time.time()
    successful_count = 0
    failed_processes = []

    robo.consulta_instrumento()

    for indice in range(start_index, max_linha):
        eta(indice=indice, max_linha=max_linha, start_time=start_time)
        print(f"\n{'⚡' * 3}🚀 EXECUTANDO PROCESSO: {numero_processo[indice]} 🚀{'⚡' * 3}".center(70, '='))
        print(f"📊 Progresso: {indice + 1}/{max_linha} ({(indice + 1) / max_linha * 100:.1f}%)")

        try:
            # Execute search
            robo.loop_de_pesquisa(
                numero_processo=numero_processo[indice],
                caminho_excel=caminho_arquivo_destino
            )

            # If successful, save progress
            successful_count += 1
            monitor.save_progress(indice, numero_processo[indice])
            print(f"✅ Processo {numero_processo[indice]} concluído com sucesso!")

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\n❌ Erro ao processar o índice {indice} ({numero_processo[indice]}): {type(e).__name__}")
            print(f"   Erro: {str(e)[:100]}")

            # Record failure
            failed_processes.append({
                'index': indice,
                'process': numero_processo[indice],
                'error': str(e)[:200],
                'error_type': type(e).__name__,
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

            # Save failed processes to file
            with open("failed_processes.json", 'w', encoding='utf-8') as f:
                json.dump(failed_processes, f, indent=4, ensure_ascii=False)

            print(f"⚠️ Falha registrada. Continuando com o próximo processo...")

            # Try to recover
            try:
                robo.consulta_instrumento()
                print("🔄 Navegação recuperada, continuando...")
            except Exception as recovery_error:
                print(f"⚠️ Falha na recuperação automática: {recovery_error}")
                print("🔄 Tentando reinicializar o navegador...")
                try:
                    robo = Robo()  # Recreate the robot instance
                    robo.consulta_instrumento()
                except:
                    print("❌ Não foi possível recuperar. Parando execução.")
                    break

            # Save progress even on failure to mark where we stopped
            monitor.save_progress(indice, numero_processo[indice])
            continue

    # Final summary
    end_time = time.time()
    tempo_total = end_time - start_time
    horas = int(tempo_total // 3600)
    minutos = int((tempo_total % 3600) // 60)
    segundos = int(tempo_total % 60)

    print("\n" + "=" * 70)
    print("📊 RESUMO FINAL".center(70))
    print("=" * 70)
    print(f"✅ Processos bem-sucedidos: {successful_count}/{max_linha - start_index}")
    print(f"❌ Processos com falha: {len(failed_processes)}")
    print(f"⏳ Tempo total de execução: {horas}h {minutos}m {segundos}s")

    if failed_processes:
        print("\n❌ Processos com falha:")
        for fail in failed_processes:
            print(f"   - Índice {fail['index']}: {fail['process']} - {fail['error_type']}")
        print(f"\n📁 Detalhes das falhas salvos em 'failed_processes.json'")

    monitor.show_summary(max_linha)
    robo.condicao_email(caminho_pasta=r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Python\webscraping\Acompanhamento")
    #send_emails_from_excel()
    print("\n👋 Execução finalizada!")


if __name__ == "__main__":
    start_time = time.time()

    try:
        main()
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"\nError occurred at line: {exc_tb.tb_lineno}")
        print(f"❌ Erro fatal na execução principal: {type(e).__name__}.\n Erro: {str(e)[:80]}")

    end_time = time.time()
    tempo_total = end_time - start_time
    horas = int(tempo_total // 3600)
    minutos = int((tempo_total % 3600) // 60)
    segundos = int(tempo_total % 60)
    print(f'⏳ Tempo de execução: {horas}h {minutos}m {segundos}s')