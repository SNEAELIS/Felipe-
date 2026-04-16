import os
import time
import re
import sys
from datetime import timedelta
from time import process_time_ns

import pandas as pd

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError
from playwright.sync_api import sync_playwright, Locator

from colorama import Fore, Style
from tqdm import tqdm

class BreakInnerLoop(Exception):
    pass


class PWRobo:
    def __init__(self, caminho_arquivo_saida: str, cdp_url: str = "http://localhost:"):
        self.port = 9222
        self.caminho_arquivo_saida = caminho_arquivo_saida
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.connect_over_cdp(f'{cdp_url}{self.port}')

        self.context = self.browser.contexts[0] if self.browser.contexts else self.browser.new_context()
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()

        # Initialize output dataframe and counter for batch saving
        self.unsaved_count = 0
        if os.path.exists(self.caminho_arquivo_saida):
            try:
                self.output_df = pd.read_excel(self.caminho_arquivo_saida)
            except Exception:
                self.output_df = pd.DataFrame()
        else:
            self.output_df = pd.DataFrame()

        # Enable resource blocking for faster performance
        #self.block_rss()

        print(f"✅ Connected to existing Chrome instance via Playwright. Connected to page: {self.page.url}")
        

    def block_rss(self):
        """Block images, stylesheets, and fonts to make browsing faster"""

        def route_handler(route, request):
            if request.resource_type in ["image", "stylesheet", "font"]:
                route.abort()
            else:
                route.continue_()

        self.page.route('**/*', route_handler)
        print("✅ Resource blocking enabled for faster performance")


    def consulta_proposta(self):
        """Navigates through the system tabs, on the landing page, until the process search page."""
        try:
            self.page.locator('xpath=//*[@id="logo"]/a').click(timeout=800)
        except PlaywrightTimeoutError:
            print('Already on the initial page of transferegov discricionárias.')
        except PlaywrightError as e:
            print(Fore.RED + f'🔄❌ Failed to reset.\nError: {type(e).__name__}\n{str(e)[:100]}')

        xpaths = ['xpath=//*[@id="menuPrincipal"]/div[1]/div[4]',
                  'xpath=//*[@id="contentMenu"]/div[1]/ul/li[6]/a']
        try:
            for xpath in xpaths:
                self.page.locator(xpath).click(timeout=10000)
        except PlaywrightError as e:
            print(Fore.RED + f'🔴📄 Instrument unavailable. \nError: {e}')
            sys.exit(1)


    def campo_pesquisa(self, numero_processo):
        """ Fill in the search field and access the desired program"""
        try:
            campo_pesquisa_locator = self.page.locator('xpath=//*[@id="consultarNumeroConvenio"]')
            campo_pesquisa_locator.fill(numero_processo)
            campo_pesquisa_locator.press('Enter')
            try:
                acessa_item_locator = self.page.locator('xpath=//*[@id="tbodyrow"]/tr/td[1]/div/a')
                acessa_item_locator.click(timeout=8000)
            except PlaywrightTimeoutError:
                print(f' Process number: {numero_processo}, not found.')
                raise BreakInnerLoop
            except PlaywrightError as e:
                print(f' Process number: {numero_processo}, not found. Error: {type(e).__name__}')
                raise BreakInnerLoop
        except PlaywrightError as e:
            print(f' Failed to insert process number in the search field. Error: {type(e).__name__}')


    def dados_proposta(self, tr: Locator) -> dict[str, str] | None:
        try:
            arvore_element = tr.locator('td.arvoreValores.arvoreExibe')

            # extract all text and parse if there is at least one element in the tree
            if arvore_element.count() > 0:
                print("Acessando árvore de valores")
                all_text = arvore_element.first.inner_text()
                lines = [line.strip() for line in all_text.split('\n') if line.strip()]
                # Pair values with their labels based on the structure
                return  {text[2]: ' '.join(text[0:2]) for line in lines for text in [line.split(maxsplit=2)]}

            labels = tr.locator('td.label').all_text_contents()
            fields = tr.locator('td.field').all_text_contents()

            all_data = {str(label).strip(): str(field).strip() for label, field in zip(labels, fields) if label.strip()}

            #if all_data.get('Modalidade') == 'Convênio':

            try:
                return all_data

            except Exception as e:
                print(f'Erro tentando processar dados da tabela. Erro: {type(e).__name__}.\nDescrição do erro: '
                      f'{str(e)[:100]}')

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Unexpected error in dados_proposta: {type(e).__name__} - {str(e)[:100]}")
            # Ensure we navigate back even on error
            self.page.locator('xpath=//*[@id="lnkConsultaAnterior"]').click(timeout=2000)


    def save_data(self):
        """Saves the buffered data to Excel."""
        if self.unsaved_count > 0:
            try:
                self.output_df.to_excel(self.caminho_arquivo_saida, index=False)
                print(f"💾 Batch saved {self.unsaved_count} records.")
                self.unsaved_count = 0
            except Exception as e:
                print(f"❌ Error saving batch: Erro: {type(e).__name__}.\nError description: {str(e)[:100]}")


    def mark_as_done(self, raw_data: list, numero_processo):
        """Safely mark row as done in the DataFrame"""

        def merge_data(raw_data: list) -> dict:
            """Merge partial dictionaries into complete records with duplicate key protection"""
            final_record = {}

            for partial_dict in raw_data:
                if not partial_dict:
                    continue

                for key, value in partial_dict.items():
                    original_key = str(key).strip()
                    new_key = original_key
                    counter = 1

                    # If key exists, find a new name (Key_1, Key_2, etc.)
                    while new_key in final_record:
                        new_key = f"{original_key}_{counter}"
                        counter += 1

                    final_record[new_key] = value

            return final_record

        def sanitize_txt(txt):
            if isinstance(txt, str):
            # Remove control characters that are not allowed in XML 1.0 (used by .xlsx)
            # Allowed: \x09 (tab), \x0A (newline), \x0D (carriage return)
            # Disallowed: \x00-\x08, \x0B, \x0C, \x0E-\x1F
                return re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F]', '', txt)
    
            return txt
       

        row_data = merge_data(raw_data)

        row_data = {k: sanitize_txt(v) for k, v in row_data.items()}

        try:
            df = self.output_df

            if not df.empty:
                # Check if process exists to update instead of append
                id_col = df.columns[0]
                mask = df[id_col].astype(str).str.strip() == str(numero_processo).strip()

                if mask.any():
                    idx = df.index[mask][0]
                    # Update existing row
                    for key, value in row_data.items():
                        if key not in df.columns:
                            df[key] = pd.NA
                        df.at[idx, key] = value
                    print(f"✅ Updated existing process: {numero_processo}")
            # Append new row
                else:
                    # Append new row
                    df_update = pd.DataFrame([row_data])
                    self.output_df = pd.concat([df, df_update], ignore_index=True)
                    print(f"✅ Appended new process: {numero_processo}")

            else:
                self.output_df = pd.DataFrame([row_data])
                print(f"✅ Appended new process: {numero_processo}")

            self.unsaved_count += 1
            
            if self.unsaved_count >= 10:
                self.save_data()

        except Exception as err:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Error in mark_as_done:{type(err).__name__}\n{str(err)[:100]}\n")
            return False


    def loop_de_pesquisa(self, numero_processo: str):
        print(f'🔍 Starting data extraction loop'.center(50, '-'), '\n')

        xpath = '//*[@id="alterar"]/div/form/table'
        try:
            self.campo_pesquisa(numero_processo=numero_processo)

            self.page.locator(xpath).wait_for(state='visible')

            # Localiza a tabela principal da aba dados
            rows = self.page.locator(xpath).locator('tr')

            raw_row_data = list()

            for _ in  range(rows.count()):
                row = rows.nth(_)
                row_data = self.dados_proposta(row)
                raw_row_data.append(row_data)

            self.mark_as_done(numero_processo=numero_processo, raw_data=raw_row_data)

            self.page.locator('xpath=//*[@id="breadcrumbs"]/a[2]').click()

            print(f"{'✨' * 3}📦 DADOS COLETADOS COM SUCESSO 📦{'✨' * 3}")
        except PlaywrightTimeoutError as t:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f'TIMEOUT {str(t)[:50]}')
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            self.consulta_proposta()
            return False
        except Exception as erro:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f'❌ Failed to try to include documents. Error:{type(erro).__name__}\n'
                  f'{str(erro)[:100]}...')
            self.consulta_proposta()
            raise BreakInnerLoop


    @staticmethod
    def extrair_dados_excel(caminho_arquivo_fonte, filter_=False):
        try:
            complete_data_frame = pd.read_excel(caminho_arquivo_fonte, dtype=str, sheet_name=None)
            sheet_names_list = list(complete_data_frame.keys())

            print("✅ Sheet Names Found:")
            for name in sheet_names_list:
                print(f"- {name}")

            data_frame = complete_data_frame['Aspar']


            if filter_:
                data_frame = data_frame[data_frame['Instrumento'].astype(str).str.contains('/',na=False)]
                data_frame = data_frame.drop_duplicates()

            print(f"✅ Loaded {len(data_frame)} rows from Excel.")
            return data_frame

        except Exception as e:
            print(f"🤷‍♂️❌ Error reading the excel file: {os.path.basename(caminho_arquivo_fonte)}.\n"
                  f"Error name: {type(e).__name__}\nError: {str(e)}")


    @staticmethod
    def fix_prop_num(numero_proposta):
        if pd.isna(numero_proposta):
            return False

        numero_proposta = str(numero_proposta)

        if '_' in numero_proposta:
            if '/' in numero_proposta:
                numero_proposta = numero_proposta.replace('_', '/')
                left, right = numero_proposta.split('/', 1)

                fixed = f"{left.zfill(6)}/{right}"

                pattern = r'^\d{6}/\d{4}'
                if re.match(pattern, fixed):
                    return fixed
                else:
                    return None

        else:
            if '/' in numero_proposta:
                left, right = numero_proposta.split('/', 1)

                fixed = f"{left.zfill(6)}/{right}"

                pattern = r'^\d{6}/\d{4}'
                if re.match(pattern, fixed):
                    return fixed
                else:
                    return None

def get_number_part(proposal):
    """Extract the number before the slash and remove leading zeros"""
    return str(proposal)


def main() -> None:
    dir_path = (r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Obras Maranhão - ASPAR-SNEAELIS.xlsx")


    output_path = (r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência "
                r"Social\Teste001\resultado_aba_dados_1-2.xlsx")

    try:
        robo = PWRobo(output_path)
        other_door = 9222#input('Enter the door you are using:\n If you are using 9222, just type enter.  ')
        robo.port = str(other_door)
    except Exception as e:
        print(f"\n‼️ Fatal error starting the robot: {e}")
        sys.exit("Stopping the program.")

    # Load DataFrame
    df = robo.extrair_dados_excel(dir_path)
    if df is None:
        print("❌ Failed to load Excel file. Exiting.")
        return
    if os.path.exists(output_path):
        # Load the existing output file
        df_existing = pd.read_excel(output_path)

        # Find completed programs (with Valor Global filled)
        if 'Valor Global' in df_existing.columns and 'Código do Instrumento' in df_existing.columns:
            finished_programs = df_existing[
                df_existing['Valor Global'].notna() &
                (df_existing['Valor Global'] != '')
                ]

            # Get list of completed proposal numbers
            to_jump = finished_programs['Código do Instrumento'].tolist()

            if to_jump:
                print(f"✅ Port {9222}: Found {len(to_jump)} already completed proposals")

                # Extract number parts for comparison
                finished_numbers = {get_number_part(p) for p in to_jump }
                source_numbers = df.iloc[:, 0].apply(get_number_part)

                # Create filter mask
                filter_mask = source_numbers.isin(finished_numbers)
                filtered_out = filter_mask.sum()

                # Apply filter
                df = df[~filter_mask]
                df = df.dropna(how='all')

                print(f"🎯 Port {9222}: Filtered out {filtered_out} already completed proposals")
            else:
                print(f"📭 Port {9222}: No completed proposals found in output file")
        else:
            print(f"⚠️ Port {9222}: Output file missing required columns")
    else:
        print(f"🆕 Port {9222}: No existing output file, processing all {len(df)} proposals")
    robo.consulta_proposta()

    start_time = time.time()

    print(f'Filtered DF has: {len(df)} lines.')

    for idx, row in tqdm(df.iterrows(), total=len(df), desc="Processing First Half", unit="prop"):
        numero_processo_temp = row.iloc[0]
        #numero_processo = robo.fix_prop_num(numero_processo_temp)
        numero_processo = numero_processo_temp

        if not numero_processo or numero_processo == 'nan' :
            continue

        tqdm.write(f"\n{'⚡' * 3}🚀 EXECUTING PROPOSAL: {numero_processo} 🚀{'⚡' * 3}".center(70, '='))

        try:
            robo.loop_de_pesquisa(numero_processo=numero_processo)

        except BreakInnerLoop:
            robo.save_data()
            sys.exit(0)
        except KeyboardInterrupt:
            print("Script stopped by user (Ctrl+C).")
            robo.save_data()
            sys.exit(1)
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            tqdm.write(f"Error occurred at line: {exc_tb.tb_lineno}")
            tqdm.write(f"❌ Failed to execute script. Error: {type(e).__name__}\n{str(e)}")
            # Save progress on error
            robo.save_data()

    robo.save_data()
    end_time = time.time()
    elapsed_time = str(timedelta(seconds=int(end_time - start_time)))
    print(f"\n🎉 Processing complete! Total execution time: {elapsed_time}")


if __name__ == "__main__":
    main()