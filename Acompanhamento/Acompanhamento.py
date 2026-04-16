import time
import os
import sys
import json
import traceback
from tabnanny import check

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
                self.chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9228")
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


    def check_situacao(self, path):
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
            print(f"📋 Colunas salvas: {list(df_novo.columns)}")

            return True

        except Exception as e:
            print(f"❌ Erro ao salvar dados no Excel: {e}")
            return False


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

            print(todos_dados)
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


    def load_progress(self):
        """Load saved progress from JSON file"""
        try:
            if os.path.exists(self.progress_file):
                with open(self.progress_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.last_index = data.get('last_index', -1)
                    self.last_process = data.get('last_process', None)
                    self.completed_processes = data.get('completed_processes', [])
                    print(
                        f"📂 Progresso carregado: Último índice = {self.last_index}, Processo = {self.last_process}")
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


    def reset_progress(self):
        """Reset all progress"""
        try:
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


    @staticmethod
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


    # Verifica as condições para mandar um e-mail para o técnico
    def condicao_email(self, numero_processo: str, caminho_pasta: str):
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
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f"⚠️Nenhum documento novo encontrado para o processo {numero_processo}.\n")
            return []


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
    caminho_arquivo_fonte = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Acompanhamento\CONTROLE DE PARCERIAS CGAP_simplificado.xlsx"

    # Caminho do arquivo .xlsx que contem os dados da ultima extração do robô
    caminho_arquivo_destino = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Acompanhamento\Result_Scrape_Acomp.xlsx"

    try:
        # Instancia um objeto da classe Robo
        robo = Robo()
        # Extrai dados de colunas específicas do Excel
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"\nError occurred at line: {exc_tb.tb_lineno}")
        print(f"\n‼️ Erro fatal ao iniciar o robô: {type(e).__name__}.\n Erro: {str(e)[:80]}")
        sys.exit("Parando o programa.")

    try:
        numero_processo = robo.extrair_dados_excel(
            caminho_arquivo_fonte=caminho_arquivo_fonte,
            busca_id='Instrumento nº',
        )
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"\nError occurred at line: {exc_tb.tb_lineno}")
        print(f"❌ Erro ao extrair dados do Excel: {type(e).__name__}.\n Erro: {str(e)[:80]}")
        sys.exit("Parando o programa.")

    # Pós-processamento dos dados para não haver erros na execução do programa
    numero_processo = robo.limpa_dados(numero_processo)

    # Inicia o processo de consulta do instrumento
    try:
        robo.consulta_instrumento()
    except Exception as e:
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"\nError occurred at line: {exc_tb.tb_lineno}")
        print(f"❌ Erro ao consultar instrumento: {type(e).__name__}.\n Erro: {str(e)[:80]}")
        sys.exit("Parando o programa.")

    max_linha = len(numero_processo)
    start_time = time.time()
    for indice in range(0, max_linha):
        eta(indice=indice, max_linha=max_linha, start_time=start_time)
        print(f"\n{'⚡' * 3}🚀 EXECUTING FILE: {numero_processo[indice]} 🚀{'⚡' * 3}".center(70, '=')
              , '\n')
        try:
            # Executa pesquisa dos termos e salva os resultados na pasta "caminho_pasta"
            robo.loop_de_pesquisa(numero_processo=numero_processo[indice],caminho_excel=caminho_arquivo_destino)

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"\nError occurred at line: {exc_tb.tb_lineno}")
            print(f"\n❌ Erro ao processar o índice {indice} ({numero_processo[indice]}): {type(e).__name__}.\n Erro: {str(e)[:80]}")
            try:
                robo.consulta_instrumento()
            except Exception as recovery_error:
                exc_type, exc_value, exc_tb = sys.exc_info()
                print(f"\nError occurred at line: {exc_tb.tb_lineno}")
                print(f"❌ Falha na recuperação: {recovery_error}")
            continue  # Continua para o próximo processo
    # Confirma se houve atualização na pasta e envia email para o técnico
    # confirma_email = list(robo.condicao_email(numero_processo=numero_processo[indice]))


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