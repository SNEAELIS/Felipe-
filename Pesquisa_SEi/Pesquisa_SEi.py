import pandas as pd
import re
import time
import os
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException, NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup


# --- Função de log ---
def registrar_log(mensagem: str):
    timestamp = datetime.now().strftime("[%Y-%m-%d %H:%M:%S]")
    with open("log.txt", "a", encoding="utf-8") as log_file:
        log_file.write(f"{timestamp} {mensagem}\n")


# --- Função resetar log ---
def resetar_log():
    with open("log.txt", "w", encoding="utf-8") as log_file:
        log_file.write("")

# --- Recupera processos já processados do log ---
def obter_processos_processados():
    if not os.path.exists("log.txt"):
        return set()
    with open("log.txt", "r", encoding="utf-8") as f:
        linhas = f.readlines()
    return {linha.split("processo: ")[1].strip() for linha in linhas if "Iniciando scraping do processo:" in linha}


# --- Conectar ao navegador existente ---
def conectar_navegador_existente(porta: int = 9222):
    try:
        print(f"Tentando conectar ao navegador na porta {porta}...")
        opcoes_chrome = Options()
        opcoes_chrome.add_experimental_option("debuggerAddress", f"127.0.0.1:{porta}")

        navegador = webdriver.Chrome(options=opcoes_chrome)


        handles = navegador.window_handles
        print(handles)
        print("Current URL:", navegador.current_url)

        print("✅ Conectado ao navegador existente com sucesso!")
        registrar_log("Conectado ao navegador com sucesso.")
        return navegador
    except WebDriverException:
        msg = "Erro ao conectar. Verifique se o Chrome está aberto com depuração."
        print("❌", msg)
        registrar_log(f"ERRO: {msg}")
        return None


# --- Ler processos válidos ---
def ler_processos_validos(caminho_excel: str, nome_coluna: str = "processo"):
    df = pd.read_excel(caminho_excel)
    print(f'Planilha carregada com {len(df)} linhas.')
    padrao = re.compile(r"^\d{5}\.\d{6}\/\d{4}-\d{2}$")  # ex: 71000.037605/2025-61
    processos_validos = df[nome_coluna].dropna().astype(str)
    processos_filtrados = [proc for proc in processos_validos if padrao.match(proc)]
    print(f"🔍 Total de processos válidos encontrados: {len(processos_filtrados)}")
    registrar_log(f"{len(processos_filtrados)} processos válidos encontrados.")
    return processos_filtrados


# --- Verificar autenticação ---
def verificar_autenticacao(navegador):
    try:
        erro_autenticacao = navegador.find_elements(By.XPATH,
            "//*[contains(text(), 'Acesso Negado') or contains(text(), 'Sessão Expirada')]")
        if erro_autenticacao:
            registrar_log("❌ Sessão expirada ou acesso negado detectado.")
            return False
        return True
    except Exception as e:
        registrar_log(f"Erro ao verificar autenticação: {e}")
        return False


# --- Extrair links e áreas do processo ---
def extrair_links_processo(navegador, numero_processo):
        try:
            msg_inicio = f"Iniciando scraping do processo: {numero_processo}"
            registrar_log(msg_inicio)
            print(f"🔄 {msg_inicio}")

            if not verificar_autenticacao(navegador):
                return [{
                    "processo": numero_processo,
                    "texto_link": "Erro: Sessão expirada ou acesso negado",
                    "SNEALIS": "",
                    "CGEALIS": "",
                    "CGFP": "",
                    "CGC": "",
                    "CGAP": ""
                }]

            navegador.switch_to.window(navegador.window_handles[0])
            wait = WebDriverWait(navegador, 10)
            campo_busca = wait.until(EC.presence_of_element_located((By.ID, "txtPesquisaRapida")))
            campo_busca.clear()
            campo_busca.send_keys(numero_processo)
            campo_busca.send_keys(Keys.ENTER)
            time.sleep(2)

            iframe = wait.until(EC.presence_of_element_located((By.ID, "ifrVisualizacao")))
            navegador.switch_to.frame(iframe)
            print("🔁 Mudou para o iframe ifrVisualizacao.")

            try:
                div_arvore = wait.until(EC.presence_of_element_located((By.ID, "divArvoreHtml")))
            except:
                print(f"⚠️ Selenium falhou. Tentando BeautifulSoup para {numero_processo}")
                registrar_log(f"⚠️ Tentando BeautifulSoup para {numero_processo}")
                iframe_content = navegador.execute_script(
                    "return document.getElementById('ifrVisualizacao').contentDocument.body.innerHTML")
                soup = BeautifulSoup(iframe_content, 'html.parser')
                div_arvore_bs = soup.find("div", {"id": "divArvoreHtml"})

                if div_arvore_bs:
                    raw_text = div_arvore_bs.get_text(separator="\n").strip()
                    return [{
                        "processo": numero_processo,
                        "texto_link": raw_text,
                        "SNEALIS": "SNEALIS" if "SNEALIS" in raw_text else "",
                        "CGEALIS": "CGEALIS" if "CGEALIS" in raw_text else "",
                        "CGFP": "CGFP" if "CGFP" in raw_text else "",
                        "CGC": "CGC" if "CGC" in raw_text else "",
                        "CGAP": "CGAP" if "CGAP" in raw_text else ""
                    }]
                else:
                    return [{
                        "processo": numero_processo,
                        "texto_link": "Erro: divArvoreHtml não encontrado (BS)",
                        "SNEALIS": "",
                        "CGEALIS": "",
                        "CGFP": "",
                        "CGC": "",
                        "CGAP": ""
                    }]

            # 🔽 FLUXO PRINCIPAL NORMAL DO SELENIUM 🔽
            try:
                div_info = div_arvore.find_element(By.ID, "divInformacao")
                raw_text = div_info.get_attribute("innerText").strip()
            except NoSuchElementException:
                raw_text = div_arvore.get_attribute("innerText").strip()

            links = div_arvore.find_elements(By.XPATH, ".//a[@href]")
            resultados = []

            if links:
                for link in links:
                    try:
                        texto = link.text.strip()
                        href = link.get_attribute("href")
                        if texto and href:
                            texto_upper = texto.upper()  # Força a comparação ser case-insensitive

                            result = {
                                "processo": numero_processo,
                                "texto_link": texto,
                                "SNEALIS": "SNEALIS" if "SNEALIS" in texto_upper else "",
                                "CGEALIS": "CGEALIS" if "CGEALIS" in texto_upper else "",
                                "CGFP": "CGFP" if "CGFP" in texto_upper else "",
                                "CGC": "CGC" if "CGC" in texto_upper else "",
                                "CGAP": "CGAP" if "CGAP" in texto_upper else ""
                            }
                            print(f"Console Log - Processo {numero_processo}, Link: {result}")
                            resultados.append(result)
                    except Exception as e:
                        registrar_log(f"Erro ao processar link em {numero_processo}: {e}")
                        continue
                print(f"✅ {len(resultados)} links encontrados para {numero_processo}")
                registrar_log(f"{len(resultados)} links encontrados para {numero_processo}.")
            else:
                print(f"ℹ️ Processo {numero_processo} sem links.")
                registrar_log(f"ℹ️ Processo {numero_processo} sem links.")
                resultados.append({
                    "processo": numero_processo,
                    "texto_link": raw_text,
                    "SNEALIS": "SNEALIS" if "SNEALIS" in raw_text else "",
                    "CGEALIS": "CGEALIS" if "CGEALIS" in raw_text else "",
                    "CGFP": "CGFP" if "CGFP" in raw_text else "",
                    "CGC": "CGC" if "CGC" in raw_text else "",
                    "CGAP": "CGAP" if "CGAP" in raw_text else ""
                })

            return resultados  # <- ESSENCIAL PARA O FLUXO NORMAL DO SELENIUM ✅

        except Exception as e:
            msg = f"❌ Erro no processo {numero_processo}: {e}"
            print(msg)
            registrar_log(msg)
            return [{
                "processo": numero_processo,
                "texto_link": f"Erro: {str(e)}",
                "SNEALIS": "",
                "CGEALIS": "",
                "CGFP": "",
                "CGC": "",
                "CGAP": ""
            }]


# --- Função principal ---
def executar_scraping():
    def eta(indice):
        elapsed_time = time.time() - start_time

        # Média por iteração
        avg_time_per_iter = elapsed_time / indice

        # Estimativa de tempo restante
        remaining_iters = max_linha - indice
        eta_seconds = remaining_iters * avg_time_per_iter

        # Formata ETA como mm:ss
        eta_minutes = int(eta_seconds // 60)
        eta_secs = int(eta_seconds % 60)

        print(
            f"\n{indice} {'>' * 10} Porcentagem concluída:"
            f" {(indice / max_linha) * 100:.2f}% | ETA: {eta_minutes:02d}:{eta_secs:02d}\n")

    reset_log = input(str('Deseja resetar o log? [Y/n]'))
    if reset_log == 'Y':
        resetar_log()

    registrar_log("=== INÍCIO DA EXECUÇÃO ===")
    navegador = conectar_navegador_existente()
    if not navegador:
        registrar_log("Execução abortada: navegador não conectado.")
        return
    arquivo_fonte = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                  r'Social\Teste001\Processo_SEi bkp\baseDados_Formalização_2024_SEi.xlsx')

    arquivo_destino = (
        r"C:\Users\felipe.rsouza\Ministério do Desenvolvimento e Assistência Social\SNEAELIS - Power "
        r"BI\Formalização\BaseDados_Formalização_2024_SEi.xlsx"
    )

    processos_todos = ler_processos_validos(arquivo_fonte)
    max_linha = len(processos_todos) + 1

    processos_processados = obter_processos_processados()

    todos_os_resultados = []
    indice = 1
    for processo in processos_todos:
        eta(indice)
        if processo in processos_processados:
            print(f"🔄 Ignorando {processo}.")
            continue

        print(f"🚀 Executando: {processo}")
        resultado = extrair_links_processo(navegador, processo)
        todos_os_resultados.extend(resultado)

        # Exporta incrementalmente
        df_parcial = pd.DataFrame(todos_os_resultados)
        df_parcial.to_excel(arquivo_destino, index=False)

        indice += 1


    registrar_log("✅ Execução finalizada com sucesso.")
    registrar_log("=== FIM DA EXECUÇÃO ===\n")


# --- Execução ---
if __name__ == "__main__":
    start_time = time.time()

    executar_scraping()