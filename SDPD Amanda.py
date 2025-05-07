from selenium import webdriver
from selenium.webdriver.common.by import By
#from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from itertools import count
import time
import pandas as pd
import os


def conectar_navegador_existente():
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"
    try:
        driver = webdriver.Chrome(options=options)
        print("✅ Conectado ao navegador existente!")
        return driver
    except Exception as erro:
        print(f"❌ Erro ao conectar ao navegador: {erro}")
        exit()

def buscar_dados(driver, numero_processo):
    try:
        # Acessar o site
        driver.get("https://www2.scdp.gov.br/novoscdp/pages/main.xhtml")
        time.sleep(0.5)

        # Clicar no menu principal
        driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/form/div/ul/li[3]/a/span[1]").click()
        driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/form/div/ul/li[3]/a/span[1]").click()
        time.sleep(0.5)

        # Clicar no submenu
        driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/form/div/ul/li[3]/ul/li[1]/a/span").click()
        time.sleep(0.5)

        # Buscar número do processo
        campo_busca = driver.find_element(By.XPATH,
                        "/html/body/div[1]/div[4]/div[4]/div[1]/div/div[1]/form/span[2]/fieldset/div/div[1]/input")
        campo_busca.clear()
        campo_busca.send_keys(numero_processo)

        # Clicar no botão de busca
        driver.find_element(By.XPATH,
                        "/html/body/div[1]/div[4]/div[4]/div[1]/div/div[1]/form/span[2]/"
                        "fieldset/div/div[3]/button").click()
        time.sleep(0.5)

        # Coletar dados
        numero_pcdp = driver.find_element(By.XPATH,
                             '//*[@id="cabecalhoViagem:numSolicitacao:numSolicitacao_text:outputText"]').text
        nome_proposto = driver.find_element(By.XPATH,
                             '//*[@id="cabecalhoViagem:nomeProposto"]').text
        tipo_viagem = driver.find_element(By.XPATH,
                             '//*[@id="cabecalhoViagem:tipoViagem:tipoViagem_text:outputText"]').text
        try:
            status_viagem = driver.find_element(By.XPATH,
                              '//*[@id="cabecalhoViagem:pnInfoPcdp"]/div[1]/div[3]/div/div[1]/a').text
        except NoSuchElementException:
            status_viagem = driver.find_element(By.XPATH,
                              '//*[@id="cabecalhoViagem:pnInfoPcdp"]/div[1]/div[3]/div/div[1]/span').text
        motivo_viagem = driver.find_element(By.XPATH,
                              '//*[@id="cabecalhoViagem:descricaoMotivoViagemText:outputText"]').text
        data_solicitacao = driver.find_element(By.XPATH,
                              '//*[@id="cabecalhoViagem:dataSolicitacao:outputText"]').text
        data_ida = driver.find_element(By.XPATH,
                              '//*[@id="cabecalhoViagem:periodoInicioViagem:outputText"]').text
        data_volta = driver.find_element(By.XPATH,
                              '//*[@id="cabecalhoViagem:periodoFimViagem:outputText"]').text
        origem = driver.find_element(By.XPATH,
                              '//*[@id="roteirosViagemSolicitacao:j_idt1099:resumoListaTrecho'
                              ':0:j_idt1106:outputText"]').text
        destino = driver.find_element(By.XPATH,
                              '//*[@id="roteirosViagemSolicitacao:j_idt1099:resumoListaTrecho:0:'
                              'j_idt1109:outputText"]').text


        return {
            "Nº da PCDP": numero_pcdp,
            "Nome do Proposto": nome_proposto,
            "Tipo de Viagem": tipo_viagem,
            "Status da Viagem": status_viagem,
            "Motivo da Viagem": motivo_viagem,
            "Data da Solicitação": data_solicitacao,
            "Data de Ida": data_ida,
            "Data de Volta": data_volta,
            "Origem": origem,
            "Destino": destino,

        }
    except NoSuchElementException:
        print(f"❌ Elemento não encontrado para o processo {numero_processo}. Continuando...")
        return None
    except Exception as erro:
        print(f"❌ Erro ao buscar dados para {numero_processo}: {erro}")
        return None


def main():
    driver = conectar_navegador_existente()
    numeros_processo = count(start=1)
    dados_coletados = []
    encerra_busca = 0

    for numero in numeros_processo:
        numero_formatado = f"{str(numero).zfill(6)}/25"
        print(f"🔍 Buscando dados para {numero_formatado}")
        try:
            dados = buscar_dados(driver, numero_formatado)
            if dados:
                dados_coletados.append(dados)
                print(f'🔍 Busca efetuada com sucesso, dados concatenados !')
                encerra_busca = 0
            elif not dados:
                encerra_busca += 1
                print(f'⏳💀  Processo não encontrado, iniciando sequencia de encerramento.\n'
                  f'Falha de busca {encerra_busca} de 3')
                if encerra_busca > 3:
                    print(f"🚫 Limite de falhas consecutivas atingido. Encerrando busca.")
                    break
        except Exception:
            print('🔍❌ Falha em realizar a busca')


    # Criar DataFrame e salvar planilha
    if dados_coletados:
        df = pd.DataFrame(dados_coletados)
        arquivo = "dados_viagens.xlsx"
        df.to_excel(arquivo, index=False)
        # Encontra o caminho do arquivo, em qual pasta está salvo
        caminho_asoluto_arquivo = os.path.abspath(arquivo)
        print(f"📁 Planilha gerada com sucesso!\nCaminho absoluto do arquivo:{caminho_asoluto_arquivo}")
    else:
        print("📁 Nenhum dado foi coletado para gerar a planilha.")

if __name__ == "__main__":
    main()
