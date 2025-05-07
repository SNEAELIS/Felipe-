#import time
from datetime import datetime
import pandas as pd
#import os
from selenium import webdriver
#from selenium.common import TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from CGC_Planilha_de_Custos import XPATH_lista

# Caminhos dos arquivos
CAMINHO_ENTRADA = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Robo TransfereGov"
CAMINHO_SAIDA = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Robo TransfereGov"


def conectar_navegador_existente():
    """Conecta ao navegador Chrome já aberto na porta 9222."""
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        print("✅ Conectado ao navegador existente!")
        return driver
    except Exception as erro:
        print(f"❌ Erro ao conectar ao navegador: {erro}")
        exit()


def ler_planilha_entrada():
    """Lê a planilha de entrada e retorna um DataFrame formatado corretamente."""
    try:
        df = pd.read_excel(CAMINHO_ENTRADA, dtype=str) # Força a leitura como string para evitar NaN inesperados
        df.columns = df.columns.str.strip()  # Remove espaços em branco dos nomes das colunas   '  ' > ''
    except Exception as erro:
        print(f"❌ Erro ao carregar a planilha de entrada: {erro}")
        exit()





    # ✅ Remover espaços extras dos nomes das coluna
    df.columns = df.columns.str.strip()
    print("📋 Colunas carregadas:", df.columns.tolist())


    colunas_esperadas = ["Instrumento nº", "Técnico", "e-mail do Técnico"]
    colunas_faltando = [col for col in colunas_esperadas if col not in df.columns]

    if colunas_faltando:
        raise ValueError(f"🚨 Erro: As colunas {colunas_faltando} não foram encontradas na planilha!")

        # ✅ Garantir que a coluna "Status" está limpa e remover espaços invisíveis
    if "Status" in df.columns:
        df["Status"] = df["Status"].str.strip().replace({"": "DESCONHECIDO"}).fillna("DESCONHECIDO")

        # ✅ Filtrar somente registros com "ATIVOS TODOS"
        df = df[df["Status"] == "ATIVOS TODOS"]

    print("📋 Registros filtrados com 'ATIVOS TODOS':", len(df))


    # ✅ Garantir que todas as colunas são strings e remover espaços extras
    df = df[colunas_esperadas].apply(lambda x: x.astype(str).str.strip())

    # ✅ Substituir apenas valores realmente vazios ou NaN
    df["Técnico"] = df["Técnico"].str.strip().replace({"": "NÃO ATRIBUÍDO"}).fillna("NÃO ATRIBUÍDO")
    df["e-mail do Técnico"] = df["e-mail do Técnico"].str.strip().replace({"": "SEM EMAIL"}).fillna("SEM EMAIL")

    # ✅ Debug: Exibir os primeiros registros para verificar se os valores estão corretos
    print("🔎 Primeiras linhas da planilha carregada:")
    print(df.head())

    return df



def salvar_dado_extracao(numero_instrumento, tecnico, email, data_upload):
    """Salva os dados extraídos na planilha de saída sem sobrescrever os registros anteriores."""
    try:
        df_saida = pd.read_excel(CAMINHO_SAIDA, dtype=str)  # Lê a planilha existente
    except FileNotFoundError:
        df_saida = pd.DataFrame(columns=["Instrumento nº", "Técnico", "e-mail do Técnico", "Data Upload"])

    novo_dado = pd.DataFrame([[numero_instrumento, tecnico, email, data_upload]],
                             columns=["Instrumento nº", "Técnico", "e-mail do Técnico", "Data Upload"])

    df_saida = pd.concat([df_saida, novo_dado], ignore_index=True)
    df_saida.to_excel(CAMINHO_SAIDA, index=False)

    print(f"✅ Dados salvos: {numero_instrumento} | {tecnico} | {email} | {data_upload}")


def automatizar_navegacao(driver, numero_instrumento):
    #Realiza a automação do site seguindo os passos indicados.
    wait = WebDriverWait(driver, 5)

    def clicar(xpath, descricao=""):
        """ Aguarda o elemento estar disponível e clica """
        try:
            elemento = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
            elemento.click()
            print(f"✔ {descricao}")
        except Exception as erro:
            print(f"⚠️ Erro ao clicar ({descricao}): {erro}")

    print("➡️ Navegando pelo sistema...")

    # Com o navegador aberto e login feito, busca a aba consultar instrumentos no menu execução
    clicar("/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]", "Acessando aba execução no menu principal")
    clicar("/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/ul/li[6]/a", "Acessando aba consultar instrumentos")

    # Inserir Número do Instrumento
    try:
        input_field = wait.until(EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[2]/td[2]/input"))) #Endereço da caixa de entrada de texto
        input_field.clear() #Limpa a caixa para evitar erro de entrada
        input_field.send_keys(numero_instrumento) #Entra o número do instrumento desejado
        clicar("/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[2]/td[2]/span/input",
               "Pesquisando instrumento")

    except Exception as erro:
        print(f"⚠️ Erro ao inserir número do instrumento {numero_instrumento}: {erro}")
        return None
    # Seleciona o instrumento desejado
    clicar("/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td/div/a", "Selecionando primeiro resultado")
    #clicar("/html/body/div[3]/div[15]/div[3]/div[1]/div/form/table/tbody/tr/td[2]/input[2]", "Visualizando data de upload")

    return capturar_data_ultimo_anexo(driver)


# Função para capturar a data do último anexo
def capturar_data_ultimo_anexo(driver):
    try:
        # Capturar tods os dados anteriores à justificativa
        anexos = driver.find_elements(By.XPATH, '//*[@id="tbodyrow"]/tr/td[3]/div')
        if not anexos:
            print("[INFO] Nenhuma data de anexo encontrada.")
            return "Sem anexos", "Nenhum anexo disponível"

        # Processar datas e retornar a mais recente
        datas = []
        for anexo in anexos:
            data_texto = anexo.text.strip()
            if data_texto:
                try:
                    datas.append(datetime.strptime(data_texto, "%d/%m/%Y"))
                except ValueError:
                    print(f"[ERRO] Data inválida encontrada: {data_texto}")
        if datas:
            data_mais_recente = max(datas).strftime("%d/%m/%Y")
            return data_mais_recente, "Sucesso"
        else:
            print("[INFO] Nenhuma data válida encontrada nos anexos.")
            return "Sem anexos válidos", "Nenhuma data válida encontrada"
    except Exception as e:
        print(f"[ERRO] Falha ao capturar dados de anexos: {e}")
        return "Erro ao capturar dados", str(e)

#captura os dados da proposta que aparecem na primeira partição da aba dados(execução >> Consultar Pré-Instrumento/Instrumento >> dados da proposta >> dados
def capturar_dados_da_proposta(driver):
    caminhos_dados = ['/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[1]/td[2]/table/tbody/tr/td[1]', ]
    try:
        dados = driver.find_elements(By.XPATH, '//*[@id="tbodyrow"]/tr/td[3]/div')
        if not anexos:
            print("[INFO] Nenhuma data de anexo encontrada.")
            return "Sem anexos", "Nenhum anexo disponível"

        # Processar datas e retornar a mais recente
        datas = []
        for anexo in anexos:
            data_texto = anexo.text.strip()
            if data_texto:
                try:
                    datas.append(datetime.strptime(data_texto, "%d/%m/%Y"))
                except ValueError:
                    print(f"[ERRO] Data inválida encontrada: {data_texto}")
        if datas:
            data_mais_recente = max(datas).strftime("%d/%m/%Y")
            return data_mais_recente, "Sucesso"
        else:
            print("[INFO] Nenhuma data válida encontrada nos anexos.")
            return "Sem anexos válidos", "Nenhuma data válida encontrada"
    except Exception as e:
        print(f"[ERRO] Falha ao capturar dados de anexos: {e}")
        return "Erro ao capturar dados", str(e)

def main():
    """Fluxo principal do código."""
    driver = conectar_navegador_existente()
    df_entrada = ler_planilha_entrada()

    print(f"📊 Total de registros a processar: {len(df_entrada)}")

    for _, linha in df_entrada.iterrows():
        numero_do_instrumento = linha["Instrumento nº"]
        print(f"\n🔎 Processando: {numero_do_instrumento}")

        automatizar_navegacao(driver, numero_do_instrumento)
        # Executa a função e armazena a data extraída
        data_upload_extraida, status = capturar_data_ultimo_anexo(driver)  # Agora retorna um valor correto

        # Passa a variável correta para salvar os dados
        salvar_dado_extracao(numero_do_instrumento, tecnico_responsavel, email_do_tecnico, data_upload_extraida)

    print("\n✅ Processamento concluído!")


if __name__ == "__main__":
    main()
