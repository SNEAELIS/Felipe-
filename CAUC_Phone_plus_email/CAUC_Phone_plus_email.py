import os.path
import sys

import pandas as pd

from datetime import datetime

from colorama import  Fore

from selenium import webdriver
from selenium.common import WebDriverException, TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains


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
            try:
                # Inicializa o driver do Chrome com as opções e o gerenciador de drivers
                self.driver = webdriver.Chrome(options=self.chrome_options)

            except Exception as e:
                print(f"Error with ChromeDriverManager: {e}")
                sys.exit()

            self.driver.switch_to.window(self.driver.window_handles[0])


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
        print(f'{'⚙️'*3}💼 INICIANDO CONSULTA DE PROPOSTAS 💼{'⚙️'*3}'.center(50, '='))
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


    # preenche o campo de pesquisa
    def campo_pesquisa(self, numero_processo):
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
        except Exception as e:
            print(f' Falha ao inserir número de processo no campo de pesquisa. Erro: {type(e).__name__}')


    # Pega os dados de email e telefone
    def busca_endereco(self, numero_processo: str):
        print('\n',f"{'🗺️' * 3}📍 BUSCANDO DADOS 📍{'🗺️' * 3}".center(50, '='))
        print()

        try:
            self.campo_pesquisa(numero_processo=numero_processo)

            # Botão detalhar
            self.webdriver_element_wait('//*[@id="form_submit"]')
            detalhar_btns = self.driver.find_elements(By.XPATH, '//*[@id="form_submit"]')

            for btn in detalhar_btns:
                btn_text = btn.get_attribute('value')
                if btn_text == 'Detalhar':
                    btn.click()
                    break
                else:
                    continue
            try:
                email = self.webdriver_element_wait('//*[@id="txtEmail"]').text
            except:
                email = 'Nenhum'

            try:
                phone = self.webdriver_element_wait('//*[@id="txtTelefone"]').text
            except:

                phone = 'Nenhum'
            try:
                self.webdriver_element_wait('//*[@id="lnkConsultaAnterior"]').click()
            except:
                print('Erro ao clicar no botão de voltar, continuando execução')
            print(f'Dados encontrados: {email}, {phone}')
            return email, phone

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Erro ao buscar endereço:{type(e).__name__}.\n Erro {str(e)[:100]}")


    # Consulta a proposta pelo atalho, sem precisar navegar para a tela principal
    def consulta_prop_breadcrumb(self):
        try:
            self.webdriver_element_wait('//*[@id="breadcrumbs"]/a[2]').click()
        except:
            print(f'Erro ao tentar usar o breadcrumbs')
            self.consulta_proposta()


    @staticmethod
    def extrair_dados_excel(caminho_arquivo_fonte):
        try:
            data_frame = pd.read_excel(caminho_arquivo_fonte,dtype=str,sheet_name=1)

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


def debug_df(df, name="df", file_path=None):
    print("\n--- DEBUG:", name, "---")
    print("Shape:", df.shape)
    print("Columns:", df.columns.tolist())
    print("DTypes:", df.dtypes)
    print("Nulls:", df.isna().sum().to_dict())
    print("Head:\n", df.head())
    excel_file = pd.ExcelFile(file_path)
    print(excel_file.sheet_names)
    print("Total sheets:", len(excel_file.sheet_names))
    print("----------------------------\n")


def main() -> None:
    try:
        # Instancia um objeto da classe Robo
        robo = Robo()
    except Exception as e:
        print(f"\n‼️ Erro fatal ao iniciar o robô: {e}")
        sys.exit("Parando o programa.")

    # Caminho do arquivo .xlsx que contem os dados necessários para rodar o robô
    caminho_arquivo_fonte = (r'C:\Users\felipe.rsouza\PycharmProjects\PythonProject\Automação '
                             r'SNEAELIS\CAUC_Phone_plus_email\termos de Fomento - Pendentes de Celebração ('
                             r'14.11).xlsx')

    output_file = (r'C:\Users\felipe.rsouza\PycharmProjects\PythonProject\Automação '
                   r'SNEAELIS\CAUC_Phone_plus_email\resultado.xlsx')

    # DataFrame do arquivo excel
    df = robo.extrair_dados_excel(caminho_arquivo_fonte=caminho_arquivo_fonte)

    if not os.path.exists(output_file):
        df_copy = df.copy()
        df_copy["Email"] = ""
        df_copy["Telefone"] = ""
        df_copy.to_excel(output_file, index=False)
        print("Inicialized Excel file with empty Email/Telefone columns.")

    df_out = pd.read_excel(output_file, dtype=str)

    print('\n,'f"{'⚡' * 3}🚀 EXECUTING FILE: {os.path.basename(caminho_arquivo_fonte)}, file length: "
          f"{len(df)} 🚀{'⚡' * 3}.".center(50, '='))
    print()

    # inicia consulta e leva até a página de busca do processo
    robo.consulta_proposta()


    for idx, row in df.iterrows():
        try:
            if pd.notna(df_out.loc[idx, 'Email']):
                print('LINHA JÁ EXECUTADA')
                continue
        except KeyError:
            print(f"Column 'Email' not found in output DataFrame.")
            raise  # this is critical, stop execution

        except IndexError:
            # Happens if df_out has fewer rows than df
            print(f"Row index {idx} is out of bounds in the output Excel file.")
            raise

        except Exception as e:
            # Catch unexpected errors comfortably
            print(f"Unexpected error at row {idx}: {e}")
            raise

        try:
            numero_processo = robo.fix_prop_num(row['Nº Proposta'])

            print('\n', f"Executing for {numero_processo}. Progress: {idx/len(df):.2%}"
                  .center(50, '-'),'\n')

            email, phone = robo.busca_endereco(numero_processo=numero_processo)

            df_out.loc[idx, 'Email'] = email
            df_out.loc[idx, 'Telefone'] = phone
            df_out.to_excel(output_file, index=False)

            robo.consulta_prop_breadcrumb()


        except KeyboardInterrupt:
            print("Script stopped by user (Ctrl+C). Exiting cleanly.")
            sys.exit(0) # Exit gracefully
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Falha ao executar script. Erro: {type(e).__name__}\n{str(e)[:100]}")
            sys.exit(0)  # Exit gracefully


if __name__ == "__main__":

     main()
'''
    for i in range(20):
        cycle_start = time.time()
        print(f"\n{'=' * 50}")
        print(f"🔄 CYCLE {i + 1}/20 started at: {datetime.now().strftime('%H:%M:%S')}")
        print(f"{'=' * 50}")

        main()

        cycle_time = time.time() - cycle_start

        print(f"\n⏱️ Cycle {i + 1} took: {cycle_time / 60:.2f} minutes")
        time.sleep(1600)
'''
