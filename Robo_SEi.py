from selenium import webdriver
from selenium.common import WebDriverException
from selenium.webdriver.common.by import By
#from selenium.webdriver.common.devtools.v85.css import CSSStyle
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
#from selenium.webdriver.common.action_chains import ActionChains
#from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time

class Robo_SEi:
    def __init__(self):
       try:
            # Inicia o  WebDriver do chrome
            # Conecta ao navegador Chrome já aberto, utilizando a porta de depuração 9222.
            self.chrome_options = webdriver.ChromeOptions()
            self. chrome_options.debugger_address = 'localhost: 9222'
            self.driver = webdriver.Chrome(service = Service(ChromeDriverManager().install()), options = self.chrome_options)
            print("✅ Conectado ao navegador existente com sucesso.")
       except WebDriverException as e:
            print(f"❌ Erro ao conectar ao navegador existente: {e}")
            self.driver = None


    def pesquisa_processo(self, numero_processo):
        # Selecion o campo de pesquica rápida, insere o número do processo e aperta ENTER
        try:
            campo_pesquisa = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.ID,
                                                                         'txtPesquisaRapida')))
            campo_pesquisa.click()
            campo_pesquisa.clear()
            campo_pesquisa.send_keys(numero_processo)
            campo_pesquisa.send_keys(Keys.ENTER)

        except Exception as e:
            print(f'❌ Falha ao enviar requisição do processo')


    def acessa_email(self):
        # Abre o popup de e-mail, caso o processo esteja concluído, reabre o processo

        try:
            iframe = self.driver.find_element(By.XPATH, '//*[@id="ifrVisualizacao"]')
            self.driver.switch_to.frame(iframe)
            self.checa_estado_processo()
            email = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                               '/html/body/div[1]/div/div/div[2]/a[10]/img')))
            email.click()

        except Exception as e:
            print(f"❌ Erro de acesso encontrado, erro: {e}")

    def checa_estado_processo(self):
        # Verifica se o processo está fechado e reabre ele

        try:
            abrir_processo = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                      '/html/body/div[1]/div/div/div[2]/a[5]/img')))

            abrir_processo.get_attribute('title')
            if abrir_processo == 'Reabrir Processo':
                abrir_processo.click()

            email = WebDriverWait(self.driver, 10).until(EC.element_to_be_clickable((By.XPATH,
                                                                     '/html/body/div[1]/div/div/div[2]/a[10]/img')))
            email.click()

        except Exception as e:
            print(f"❌ Erro de acesso encontrado, erro: {e}")

    def acessa_pop_up(self):
        self.dr
        aba_principal = self.driver.current_window_handle
        todas_abas = self.driver.window_handles
        janela_pop_up = ''
        try:
            # Salva o caminho da página principal como referência
            for aba in todas_abas:
                if aba != aba_principal:
                    # Atribui um endereço para a aba de pop up
                    janela_pop_up = aba
                    break
            # Transiciona para o pop up, para que se torne a janela ativa
            self.driver.switch_to.window(janela_pop_up)
        except Exception as e:
            print(f"❌ Erro de acesso ao pop up, erro: {e}")

    def manda_email(self, email, assunto, entidade, anexo: list = None ):
        # Preenche os campos do pop up de e-mail e coloca os anexos
        try:
            destinatario = self.driver.find_element(By.XPATH,
                                                    '/html/body/div[1]/div/div/form[1]/div[4]/p/div/ul/li/input')
            destinatario.click()
            destinatario.send_keys(email)

            campo_assunto = self.driver.find_element(By.XPATH,
                                                    '/html/body/div[1]/div/div/form[1]/div[6]/input[1]')
            campo_assunto.click()
            campo_assunto.send_keys(assunto)

            mensagem = self.driver.find_element(By.XPATH,
                                                    '/html/body/div[1]/div/div/form[1]/div[6]/textarea')
            mensagem.click()
            mensagem.send_keys(f'Texto de mensagem para {entidade}')
            # Verifica quantos anexos serão enviados
            if anexo is None:
                anexo = []
            for i in len(anexo):
                campo_anexo = self.driver.find_element(By.XPATH,
                                                        '/html/body/div[1]/div/div/form[2]/div/input')
                campo_anexo.click()
                campo_anexo.send_keys('Caminho para o arquivo a ser anexado')
                campo_anexo.send_keys(Keys.ENTER)
            self.driver.find_element(By.XPATH,
                      '/html/body/div[1]/div/div/div[3]/button[1]').click()

        except Exception as e:
            print(f"❌ Erro de envio encontrado, erro: {e}")

    def ler_contatos_excel(self, file_path):
        # Busca os dados dentro da planilha, com
        dados_processo = pd.read_excel(file_path, dtype = str)
        cabecalhos_necessarios = dados_processo.columns.tolist()
        cabecalho_faltando = [chave for chave in cabecalhos_necessarios if chave not in dados_processo]
        anexo = [r"C:\Users\felipe.rsouza\Documents\Robo TransfereGov\Consultar Pré-InstrumentoInstrumento",
        r"C:\Users\felipe.rsouza\Documents\Robo TransfereGov\test SEi"]

        if cabecalho_faltando:
            raise ValueError(f'❌ Faltando o(s) cabeçalho(s):  {', '.join(cabecalho_faltando)}')

        try:
            for indice, linha in dados_processo.iterrows():  # Assume que a primeira linha e um cabeçalho
                numero_processo = linha['Processo']  # Busca o número do processo
                email = linha['email']  # Busca o destinatário da mensagem 'Texto padrão do email aqui'
                entidade = linha['Entidade']

                self.pesquisa_processo(numero_processo)
                self.acessa_email()
                self.manda_email(email, 'Teste', entidade, anexo)

        except Exception as e:
            print(f"❌ Erro de leitura encontrado, erro: {e}")

def main():
    robo = Robo_SEi()
    return robo.ler_contatos_excel(r'C:\Users\felipe.rsouza\Documents\Robo TransfereGov\test SEi.xlsx')



if __name__ == "__main__":
    main()
