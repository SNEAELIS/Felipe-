import os
import time
import re
import sys
import traceback

import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

from colorama import Fore, Style
from tqdm import tqdm


class BreakInnerLoop(Exception):
    pass


class PWRobo:
    def __init__(self, caminho_arquivo_saida: str, debugger_address: str = "localhost:9226"):
        self.caminho_arquivo_saida = caminho_arquivo_saida

        chrome_options = Options()
        chrome_options.add_experimental_option("debuggerAddress", debugger_address)

        try:
            self.driver = webdriver.Chrome(options=chrome_options) # type: ignore
            self.wait = WebDriverWait(self.driver, 10)
        except Exception as e:
            # Row indicator for connection failure
            _, _, tb = sys.exc_info()
            print(f"{Fore.RED}❌ Connection failed on Line {tb.tb_lineno}{Style.RESET_ALL}") # type: ignore
            print(f'Err: {type(e).__name__}\nErr msg: {str(e)[:100]}')

            raise

        self.unsaved_count = 0
        if os.path.exists(self.caminho_arquivo_saida):
            try:
                self.output_df = pd.read_excel(self.caminho_arquivo_saida)
            except Exception as e:
                print(f'Err: {type(e).__name__}\nErr msg: {str(e)[:100]}')
                self.output_df = pd.DataFrame()
        else:
            self.output_df = pd.DataFrame()

        print(f"✅ Connected to Chrome. Current URL: {self.driver.current_url}")

    def safe_click(self, xpath, timeout=10):
        element = WebDriverWait(self.driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        element.click()

    def consulta_proposta(self):
        try:
            logo = WebDriverWait(self.driver, 2).until(
                EC.element_to_be_clickable((By.XPATH, '//*[@id="logo"]/a'))
            )
            logo.click()
        except TimeoutException:
            print('Already on initial page.')
        except Exception as e:
            _, _, tb = sys.exc_info()
            print(f"{Fore.RED}🔄❌ Reset failed on Line {tb.tb_lineno}{Style.RESET_ALL}") # type: ignore
            print(f'Err: {type(e).__name__}\nErr msg: {str(e)[:100]}')

        try:
            self.safe_click('//*[@id="menuPrincipal"]/div[1]/div[3]')
            self.safe_click('//*[@id="contentMenu"]/div[1]/ul/li[2]/a')
        except Exception as e:
            _, _, tb = sys.exc_info()
            print(f"{Fore.RED}🔴📄 Menu navigation failed on Line {tb.tb_lineno}{Style.RESET_ALL}") # type: ignore
            print(f'Err: {type(e).__name__}\nErr msg: {str(e)[:100]}')
            sys.exit(1)

    def acomp_fisc_tab(self):
        try:
            self.safe_click("//div[contains(@class, 'button menu') and contains(text(), 'Acomp. e Fiscalização')]")
            self.safe_click("//div[contains(@class, 'border')]//a[contains(text(), 'Esclarecimentos')]")
            print(f'\n{"<" * 6} SUCESSO — Aba "Esclarecimento" acessada {">" * 6}'.center(80))
        except Exception as e:
            _, _, tb = sys.exc_info()
            print(f"{Fore.RED}❌ Error in acomp_fisc_tab on Line {tb.tb_lineno}{Style.RESET_ALL}") # type: ignore
            print(f'Err: {type(e).__name__}\nErr msg: {str(e)[:100]}')
            sys.exit(1)

    def campo_pesquisa(self, numero_processo):
        try:
            campo = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="consultarNumeroProposta"]')))
            campo.clear()
            campo.send_keys(numero_processo)
            campo.send_keys(Keys.ENTER)

            try:
                self.safe_click('//*[@id="tbodyrow"]/tr/td[1]/div/a', timeout=8)
            except TimeoutException:
                print(f' Process number: {numero_processo}, not found.')
                raise BreakInnerLoop
        except BreakInnerLoop:
            raise
        except Exception as e:
            _, _, tb = sys.exc_info()
            print(f"{Fore.RED}❌ Search failed on Line {tb.tb_lineno}{Style.RESET_ALL}") # type: ignore
            print(f'Err: {type(e).__name__}\nErr msg: {str(e)[:100]}')

    def dados_detalhamento(self) -> dict:
        tentativas = 0
        max_tentativas = 3
        
        while tentativas < max_tentativas:
            try:
                # 1. Wait specifically for the labels to appear
                self.wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "td.label")))
                
                # 2. Extract
                labels = [el.get_attribute("innerText").strip() for el in self.driver.find_elements(By.CSS_SELECTOR, "td.label")]
                fields = [el.get_attribute("innerText").strip() for el in self.driver.find_elements(By.CSS_SELECTOR, "td.field")]
                
                data = {str(l): str(f) for l, f in zip(labels, fields) if l}

                # 3. Check if empty (Retry if page didn't load content)
                if not data:
                    print(f"{Fore.YELLOW}⚠️ Page loaded but dictionary is empty. Retrying {tentativas+1}/{max_tentativas}...{Style.RESET_ALL}")
                    tentativas += 1
                    time.sleep(1.5)
                    continue

                # 5. Click Voltar
                self.driver.find_element(By.CSS_SELECTOR, 'input[value="Voltar"]').click()
                return data

            except (StaleElementReferenceException, TimeoutException) as e:
                tentativas += 1
                print(f"{Fore.YELLOW}🔄 {type(e).__name__} on attempt {tentativas}. Retrying...{Style.RESET_ALL}")
                time.sleep(1)
            
            except Exception as e:
                _, _, tb = sys.exc_info()
                print(f"{Fore.RED}❌ Hard Error on Line {tb.tb_lineno}: {type(e).__name__}{Style.RESET_ALL}") # type: ignore
                break 

        return {} # Only returns empty after 3 failed attempts       
         
    def save_data(self):
        if self.unsaved_count > 0:
            try:
                self.output_df.to_excel(self.caminho_arquivo_saida, index=False)
                print(f"💾 Batch saved {self.unsaved_count} records.")
                self.unsaved_count = 0
            except Exception as e:
                _, _, tb = sys.exc_info()
                print(f"❌ Save failed on Line {tb.tb_lineno}") # type: ignore
                print(f'Err: {type(e).__name__}\nErr msg: {str(e)[:100]}')

    def search_loop_back(self):
        try:
            self.safe_click("//div[contains(@class, 'button menu') and contains(text(), 'Propostas')]")
            self.safe_click("//div[contains(@class, 'border')]//a[contains(text(), 'Consultar Propostas')]")
        except Exception as e:
            _, _, tb = sys.exc_info()
            print(f"❌ Error in search_loop_back on Line {tb.tb_lineno}") # type: ignore
            print(f'Err: {type(e).__name__}\nErr msg: {str(e)[:100]}')

    def mark_as_done(self, raw_data_list: list, numero_processo):
        def sanitize_txt(txt):
            if isinstance(txt, str):
                return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', txt)
            return str(txt) if txt is not None else ""

        try:
            # 1. Prepare the new rows
            new_rows_list = []
            for data_item in raw_data_list:
                if data_item:
                    # Ensure primary key is in the dict and everything is sanitized
                    sanitized_row = {str(k): sanitize_txt(v) for k, v in data_item.items()}
                    new_rows_list.append(sanitized_row)

            print(len(new_rows_list))
            if not new_rows_list:
                return # Nothing to add

            # 2. Create a DataFrame from the new data
            new_data_df = pd.DataFrame(new_rows_list).astype(object)

            # 3. Append to the main DataFrame
            if self.output_df.empty:
                self.output_df = new_data_df
            else:
                # Ensure existing DF is object type to prevent float64 errors
                self.output_df = self.output_df.astype(object)
                self.output_df = pd.concat([self.output_df, new_data_df], ignore_index=True)

            print(f"✅ Added {len(new_rows_list)} rows for process: {numero_processo}")

            # 4. Batch saving logic
            self.unsaved_count += len(new_rows_list)
            if self.unsaved_count >= 10: 
                self.save_data()

        except Exception as e:
            _, _, tb = sys.exc_info()
            print(f"{Fore.RED}❌ DataFrame append failed on Line {tb.tb_lineno}") # type: ignore
            print(f'Err: {type(e).__name__}\nErr msg: {str(e)[:100]}{Style.RESET_ALL}')

    def click_page(self, page_number: int):
        selectors = [f"//a[@title='Vá para a pág {page_number}']", f"//a[text()='{page_number}']"]
        for sel in selectors:
            try:
                el = self.driver.find_element(By.XPATH, sel)
                el.click()
                return True
            except:
                continue
        return False

    def loop_de_pesquisa(self, numero_processo: str):
        print(f'🔍 Processing: {numero_processo}'.center(50, '-'))
        try:
            self.campo_pesquisa(numero_processo)
            self.acomp_fisc_tab()

            WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.ID, "esclarecimentos")))

            raw_row_data = []
            process_data = {'Número da Proposta': numero_processo}
            rows_count = len(self.driver.find_elements(By.CSS_SELECTOR, '#esclarecimentos tbody tr'))

            if rows_count > 0:
                WebDriverWait(self.driver, 5).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="esclarecimentos"]/span[1]')))
                pg_info = self.driver.find_element(By.XPATH, '//*[@id="esclarecimentos"]/span[1]').text
                match = re.search(r'\((\d+)', pg_info)
                total_det = str(match.group(1)) if match else 0
                print(f'{total_det}\n\n{match}')
                row_data = {'Número de detalhamentos': total_det}
                raw_row_data.append({**process_data, **row_data})

                self.mark_as_done(raw_row_data, numero_processo)
                self.search_loop_back()
                print(f"✅ Data collected for {numero_processo}")
            else:
                self.mark_as_done([process_data], numero_processo)
                self.search_loop_back()

        except Exception as e:
            _, _, tb = sys.exc_info()
            print(f"{Fore.RED}❌ Loop crash on Line {tb.tb_lineno}{Style.RESET_ALL}") # type: ignore
            print(f'Err: {type(e).__name__}\nErr msg: {str(e)[:100]}')
            self.consulta_proposta()
            raise BreakInnerLoop


    @staticmethod
    def get_number_part(proposal):
        """Extract the number before the slash and remove leading zeros"""
        if '/' in proposal:
            return proposal.split('/')[0].lstrip('0') or '0'
        return proposal.lstrip('0') or '0'


    @staticmethod
    def extrair_dados_excel(caminho, filter_=False):
        df = pd.read_excel(caminho, dtype=str)
        if filter_:
            df = df[df['Nº Proposta'].astype(str).str.contains('/', na=False)].drop_duplicates()
        return df

    @staticmethod
    def fix_prop_num(numero_proposta):
        if pd.isna(numero_proposta):
            return False

        pattern = r'^\d{6}/\d{4}'
        prop_str = str(numero_proposta).strip()

        if '_' in numero_proposta:
            numero_proposta_fixed = numero_proposta.replace('_', '/')
            return numero_proposta_fixed

        elif re.findall(pattern, numero_proposta):
            return numero_proposta
        else:
            if '/' in prop_str:
                parts = prop_str.split('/')
                if len(parts) == 2:
                    first_digits = re.sub(r'\D', '', parts[0])
                    second_digits = re.sub(r'\D', '', parts[1])

                    if first_digits and second_digits:
                        first_padded = f"{int(first_digits):06d}"
                        return f"{first_padded}/{second_digits}"



def main():
    output_path = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\conta_detalhamento_1-2.xlsx"
    dir_path = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Propostas_Extraidas_filtered.xlsx"

    try:
        robo = PWRobo(output_path)
    except Exception as e:
        _, _, tb = sys.exc_info() 
        print(f'Err: {type(e).__name__}\nErr msg: {str(e)[:100]}')
        sys.exit(f"Fatal Startup Error on Line {tb.tb_lineno}") # type: ignore

    df = robo.extrair_dados_excel(dir_path, filter_=True)
    robo.consulta_proposta()
     # Checks which program has been scraped alraady

    if os.path.exists(output_path):
        jump_df = robo.extrair_dados_excel(output_path)
        finished_programs = jump_df[jump_df['Número da Proposta'].notna() & (jump_df['Número da Proposta'] != '')]
        # Remove finished programs from the jump list to avoid skipping them in case they were marked as done by mistake
        to_jump = set(finished_programs['Número da Proposta'].tolist())
        # Create sets of number parts
        finished_numbers = {robo.get_number_part(p) for p in to_jump if len(p.split('/')) <= 6}
        source_numbers = df.iloc[:, 0].apply(robo.get_number_part)
        # Create mask based on number part match
        filter_mask = source_numbers.isin(finished_numbers)
        df = df[~filter_mask]

    slice_ = len(df) // 3
    str_idx = (1 - 1) * slice_
    end_idx = 1 * slice_

    for idx, row in tqdm(df[str_idx: end_idx].iterrows(), total=slice_, unit='prop'):
        numero_processo = robo.fix_prop_num(row.iloc[0])
        if not numero_processo: continue

        try:
            robo.loop_de_pesquisa(numero_processo)
        except BreakInnerLoop:
            robo.save_data()
            continue # Instead of break, try next proposal if one fails
        except KeyboardInterrupt:
            robo.save_data()
            sys.exit()

    robo.save_data()


if __name__ == "__main__":
    try:
        main()
    finally:
        print("\nPress Enter to exit...")
        input()