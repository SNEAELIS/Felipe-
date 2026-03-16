import shutil
import re
import time
import os

import pandas as pd

from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import WebDriverException, StaleElementReferenceException
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


# --- Transforma os números de processo para o formato padrão ---
def formato_padrao(num_sei: str):
    num_sei = str(num_sei).strip()

    padrao = re.compile(r"^\d{5}\.\d{6}\/\d{4}-\d{2}$")  # ex: 71000.037605/2025-61
    match_ = padrao.match(num_sei)

    if match_:
        return num_sei

    # Remove any existing separators (., /, -) to get clean number
    num_sei_limpo = re.sub(r'[./-]', '', num_sei)

    if len(num_sei_limpo) < 17:
        return ''
    else:
        part1 = num_sei_limpo[:5]  # first 5 digits
        part2 = num_sei_limpo[5:11]  # next 6 digits
        part3 = num_sei_limpo[11:15]  # next 4 digits
        part4 = num_sei_limpo[15:17]  # last 2 digits

        # Format according to pattern
        return f"{part1}.{part2}/{part3}-{part4}"


# --- Ler processos válidos ---
def ler_processos_validos(caminho_excel: str, nome_coluna: str = "processo") -> list:
    df = pd.read_excel(caminho_excel)
    print(f'Planilha carregada com {len(df)} linhas.')

    padrao = re.compile(r"^\d{5}\.\d{6}\/\d{4}-\d{2}$")  # ex: 71000.037605/2025-61
    processos_validos = df[nome_coluna].drop_duplicates().dropna().astype(str).apply(formato_padrao)
    print(processos_validos)

    processos_filtrados = [proc for proc in processos_validos if padrao.match(proc)]

    print(f"🔍 Total de processos válidos encontrados: {len(processos_filtrados)}")
    registrar_log(f"{len(processos_filtrados)} processos válidos encontrados.")

    return processos_filtrados


# --- Extrair links e áreas do processo ---
def extrair_links_processo(navegador, numero_processo):
    in_case_empty = [{
        "processo": numero_processo,
        "texto_link": "Erro: Sessão expirada ou acesso negado",
        "SNEALIS": "",
        "CGEALIS": "",
        "CGFP": "",
        "CGC": "",
        "CGAP": ""
    }]

    try:
        msg_inicio = f"Iniciando scraping do processo: {numero_processo}"
        registrar_log(msg_inicio)
        print(f"🔄 {msg_inicio}")

        navegador.switch_to.default_content()
        wait = WebDriverWait(navegador, 7)
        campo_busca = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="txtPesquisaRapida"]')))

        campo_busca.clear()
        campo_busca.send_keys(numero_processo)
        campo_busca.send_keys(Keys.ENTER)

        # Verificar se apareceu mensagem de "sem resultados"
        try:
            div_sem_resultado = WebDriverWait(navegador, 7).until(EC.presence_of_element_located((By.CLASS_NAME,
                                                                                                  "pesquisaSemResultado")))

            if div_sem_resultado and div_sem_resultado.is_displayed():
                print(f"⚠️  Processo {numero_processo}: Nenhum resultado encontrado")
                in_case_empty[0]["texto_link"] = "Nenhum resultado encontrado"
                return in_case_empty


        except StaleElementReferenceException:
            # Element became stale; re-find it
            div_sem_resultado = navegador.find_element(By.CLASS_NAME, "pesquisaSemResultado")

            if div_sem_resultado and div_sem_resultado.is_displayed():
                print(f"⚠️  Processo {numero_processo}: Nenhum resultado encontrado")
                in_case_empty[0]["texto_link"] = "Nenhum resultado encontrado"
                return in_case_empty
        except Exception:
            # Se não encontrar a div de sem resultados, continua normalmente
            pass

        confere_iframe(navegador=navegador, iframe='ifrConteudoVisualizacao')

        confere_iframe(navegador=navegador, iframe='ifrVisualizacao')

        print("🔁 Mudou para o iframe ifrVisualizacao.")

        try:
            # 2. Espera e localiza o container principal
            div_arvore = wait.until(EC.presence_of_element_located((By.ID, "divArvoreInformacao")))

            # Pegamos o texto bruto para o caso de não encontrarmos links individuais
            raw_text_total = div_arvore.get_attribute("innerText").strip()

            # 3. Busca todos os elementos <a> dentro da div
            elementos_a = div_arvore.find_elements(By.TAG_NAME, "a")

            resultados = []

            if elementos_a:
                for el in elementos_a:
                    texto = el.text.strip()
                    # SEI usa 'title' para mostrar o nome completo da unidade/pessoa
                    descricao = el.get_attribute("title") or ""

                    if texto:
                        texto_para_busca = f"{texto} {descricao}".upper()

                        # Criamos o dicionário de resultado para cada link encontrado
                        resultados.append({
                            "processo": numero_processo,
                            "texto_link": f"{texto} ({descricao})".strip(),
                            "SNEALIS": "SNEALIS" if "SNEALIS" in texto_para_busca else "",
                            "CGEALIS": "CGEALIS" if "CGEALIS" in texto_para_busca else "",
                            "CGFP": "CGFP" if "CGFP" in texto_para_busca else "",
                            "CGC": "CGC" if "CGC" in texto_para_busca else "",
                            "CGAP": "CGAP" if "CGAP" in texto_para_busca else ""
                        })

            # 4. Fallback: Se não encontrou links, mas existe texto na div
            if not resultados and raw_text_total:
                raw_upper = raw_text_total.upper()
                resultados.append({
                    "processo": numero_processo,
                    "texto_link": raw_text_total,
                    "SNEALIS": "SNEALIS" if "SNEALIS" in raw_upper else "",
                    "CGEALIS": "CGEALIS" if "CGEALIS" in raw_upper else "",
                    "CGFP": "CGFP" if "CGFP" in raw_upper else "",
                    "CGC": "CGC" if "CGC" in raw_upper else "",
                    "CGAP": "CGAP" if "CGAP" in raw_upper else ""
                })

            print(f"✅ Processado: {numero_processo} | Encontrados: {len(resultados)} itens.")
            return resultados

        except Exception as e:
            registrar_log(f"❌ Erro crítico ao extrair links do processo {numero_processo}.\nErr:"
                          f"{type(e).__name__}.\nMSG: {str(e)[:100]}")
            return []

        navegador.switch_to.default_content()

        return resultados  # <- ESSENCIAL PARA O FLUXO NORMAL DO SELENIUM ✅

    except Exception as e:
        import sys
        msg = f"❌ Erro no processo {numero_processo}"
        exc_type, exc_value, exc_tb = sys.exc_info()
        print(f"❌Error occurred at line: {exc_tb.tb_lineno}")
        print(f"{msg}: {type(e).__name__} - {str(e)[:100]}")
        registrar_log(msg)
        navegador.switch_to.default_content()

        return [{
            "processo": numero_processo,
            "texto_link": f"Erro: {str(e)}",
            "SNEALIS": "",
            "CGEALIS": "",
            "CGFP": "",
            "CGC": "",
            "CGAP": ""
        }]


# --- First, check all iframes on the page ---
def confere_iframe(navegador, iframe: str, dbg: bool = False):
    iframes = navegador.find_elements(By.TAG_NAME, "iframe")
    if dbg:
        print(f"Total iframes found: {len(iframes)}")
        for i, frame in enumerate(iframes):
            print(
                f"Iframe {i}: ID={frame.get_attribute('id')}, Name={frame.get_attribute('name')}, Src={frame.get_attribute('src')}")

    # Try to locate your specific iframe
    try:
        # Option 1: Wait for iframe and switch directly
        wait = WebDriverWait(navegador, 10)
        iframe = wait.until(EC.presence_of_element_located((By.ID, iframe)))
        navegador.switch_to.frame(iframe)
    except Exception as e:
        print(f"Error finding iframe: {e}\nError name: {type(e).__name__}\nErr: {str(e)[:100]}")


# --- Deleta todos os dados no arquivo de saída ---
def delete_destiny_data(arquivo_destino, create_backup=True):
    """
    Delete all data from the destiny file and create empty file with headers

    Args:
        arquivo_destino: Path to the destiny file
        create_backup: Whether to create a backup before deleting

    Returns:
        bool: True if successful, False otherwise
    """
    # Define the headers
    headers = ['processo', 'texto_link', 'SNEALIS', 'CGEALIS', 'CGFP', 'CGC', 'CGAP']

    if not os.path.exists(arquivo_destino):
        print(f"📁 File not found: {arquivo_destino}")
        print(f"📝 Creating new file with headers...")
    else:
        try:
            # Create backup if requested
            if create_backup:
                backup_path = arquivo_destino.replace('.xlsx',
                                                      f'_backup.xlsx')
                shutil.copy2(arquivo_destino, backup_path)
                print(f"💾 Backup created: {backup_path}")
        except Exception as e:
            print(f"⚠️ Could not create backup: {type(e).__name__}.\nMSG: {str(e)[:100]}")

    try:
        # Create empty DataFrame with headers
        empty_df = pd.DataFrame(columns=headers)

        # Save to Excel
        empty_df.to_excel(arquivo_destino, index=False)
        print(f"🗑️ All data deleted from {arquivo_destino}")
        print(f"📋 Headers maintained: {', '.join(headers)}")
        return True

    except Exception as e:
        print(f"❌ Error deleting data: {type(e).__name__}.\nMSG: {str(e)[:100]}")
        return False


# --- Filtra todos os processos já pesquisados com base no arquivo de saída ---
def obter_processos_processados(arquivo_destino, process_column='processo'):
    """
    Get list of already processed processes from destiny file

    Args:
        arquivo_destino: Path to the destiny file
        process_column: Name of the column containing process numbers

    Returns:
        set: Set of processed process numbers
    """
    df_destino = pd.read_excel(arquivo_destino, dtype=str)

    if len(df_destino) == 0:
        print("📁 Destiny file is empty or doesn't exist. No processes to filter.")
        return set()

    # Get unique processes from destiny file
    processed = set(df_destino[process_column].dropna().unique())
    print(f"📊 Found {len(processed)} already processed processes in destiny file")
    return processed


# --- Concatena os dados no arquivo exceu sem risco de destruição de dados ---
def append_to_excel_safe(df_new_data, arquivo_destino, make_backup=True):
    """
    Safely append new data to existing Excel file without overwriting

    Args:
        df_new_data: DataFrame with new data to append
        arquivo_destino: Path to the destiny file
        make_backup: Whether to create a backup before writing

    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Create backup if requested and file exists
        if make_backup and os.path.exists(arquivo_destino):
            backup_path = arquivo_destino.replace('.xlsx', f'_backup_.xlsx')
            shutil.copy2(arquivo_destino, backup_path)
            print(f"💾 Backup created: {backup_path}")

        # Read existing data DIRECTLY from Excel
        if os.path.exists(arquivo_destino):
            try:
                df_existing = pd.read_excel(arquivo_destino)
                print(f"📖 Read {len(df_existing)} existing rows from {arquivo_destino}")
            except Exception as e:
                print(f"⚠️ Could not read existing file: {type(e).__name__}.\nMSG: {str(e)[:100]}")
                df_existing = pd.DataFrame()
        else:
            df_existing = pd.DataFrame()
            print(f"📁 File doesn't exist yet, will create new")

        # Ensure df_new_data is a DataFrame
        if not isinstance(df_new_data, pd.DataFrame):
            try:
                df_new_data = pd.DataFrame(df_new_data)
                print(f"🔄 Converted new data to DataFrame")
            except:
                print(f"❌ Could not convert new data to DataFrame")
                return False

        # Combine existing and new data
        if not df_existing.empty:
            # Make sure columns align (use existing columns as reference)
            if set(df_new_data.columns) != set(df_existing.columns):
                print(f"⚠️ Column mismatch. Aligning new data to existing columns...")
                print(f"   Existing columns: {df_existing.columns.tolist()}")
                print(f"   New data columns: {df_new_data.columns.tolist()}")
                # Reindex new data to match existing columns
                df_new_data = df_new_data.reindex(columns=df_existing.columns)

            df_combined = pd.concat([df_existing, df_new_data], ignore_index=True)
            print(f"🔗 Combined {len(df_existing)} existing + {len(df_new_data)} new rows")
        else:
            df_combined = df_new_data
            print(f"📝 Using only new data ({len(df_new_data)} rows)")

        # Remove duplicates if needed (optional)
        before_dedup = len(df_combined)
        df_combined = df_combined.drop_duplicates()
        after_dedup = len(df_combined)
        if before_dedup > after_dedup:
            print(f"🗑️ Removed {before_dedup - after_dedup} duplicate rows")

        # Save combined data
        df_combined.to_excel(arquivo_destino, index=False)
        print(f"✅ Data saved to {arquivo_destino}")
        print(f"   Total rows: {len(df_combined)} (Added {len(df_new_data)} new rows)")

        return True

    except Exception as e:
        print(f"❌ Error saving to Excel: {type(e).__name__}")
        print(f"   Message: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

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
            f" {(indice / max_linha // 2) * 100:.2f}% | ETA: {eta_minutes:02d}:{eta_secs:02d}\n")

    resetar_log()
    arquivo_fonte = (r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência "
                     r"Social\Teste001\propostas_SEi.xlsx")
    arquivo_destino = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\SNEAELIS - webscraping\Consulta_SEi\Consultas parciais\consulta_direcao_sei_2-2.xlsx"

    reset_df = input(str('Deseja resetar o Data_Frame? [Y/n]'))
    if reset_df == 'Y':
        delete_destiny_data(arquivo_destino)

    registrar_log("=== INÍCIO DA EXECUÇÃO ===")

    other_door = input('Enter the door you are using:\n If you are using 9222, just type enter.  ')

    if isinstance(other_door, str) and other_door.isdigit():
        navegador = conectar_navegador_existente(int(other_door))
    else:
        navegador = conectar_navegador_existente()

    if not navegador:
        registrar_log("Execução abortada: navegador não conectado.")
        return

    processos_todos = ler_processos_validos(arquivo_fonte)
    max_linha = len(processos_todos) + 1

    processos_processados = obter_processos_processados(arquivo_destino=arquivo_destino)
    processos_processados.add('71000.063323/2024-38')

    todos_os_resultados = []
    indice = 1

    metade = len(processos_todos) // 2

    for processo in processos_todos[metade: ]:
        eta(indice)
        if processo in processos_processados:
            print(f"🔄 Ignorando {processo}.")
            continue

        print(f"🚀 Executando: {processo}")
        resultado = extrair_links_processo(navegador, processo)
        todos_os_resultados.extend(resultado)

        indice += 1

        # Exporta a cada 10 processos OU no último processo
        if indice % 10 == 0 or indice == len(processos_todos):
            print(f"💾 Salvando checkpoint após {indice} processos...")
            # Append to main file incrementally
            df_parcial = pd.DataFrame(todos_os_resultados)  # Just the last batch

            # Keep rows where 'texto_link' contains '/' AND does NOT contain '.'
            df_cleaned = df_parcial[df_parcial['texto_link'].str.contains('/') & ~df_parcial['texto_link'].str.contains(r'\.')]

            append_to_excel_safe(df_cleaned, arquivo_destino, make_backup=False)

            # Clear results to free memory (optional)
            todos_os_resultados = []

    # Exportação final garantida (caso o último grupo não seja múltiplo de 10)
    print(f"💾 Salvando dados finais...")

    # Append to main file
    df_final = pd.DataFrame(todos_os_resultados)

    df_cleaned = df_final[df_final['texto_link'].str.contains('/') & ~df_final['texto_link'].str.contains(r'\.')]

    append_to_excel_safe(df_cleaned, arquivo_destino, make_backup=True)

    registrar_log("✅ Execução finalizada com sucesso.")
    registrar_log("=== FIM DA EXECUÇÃO ===\n")


# --- Execução ---
if __name__ == "__main__":
    start_time = time.time()

    executar_scraping()