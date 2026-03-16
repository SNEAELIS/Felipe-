import time
import re
import sys
import os

import pandas as pd

from pandas import ExcelWriter

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError
from playwright.sync_api import sync_playwright


class PWRobo:
    def __init__(self, cdp_url: str = "http://localhost:9222"):
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.connect_over_cdp(cdp_url)
        self.context = self.browser.contexts[0] if self.browser.contexts else self.browser.new_context()
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()

        # Enable resource blocking for faster performance
        # self.block_rss()

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


    def search(self, numero_processo):
        try:
            campo_pesquisa_locator = self.page.locator('xpath=//*[@id="txtPesquisaRapida"]')
            campo_pesquisa_locator.fill(numero_processo)
            campo_pesquisa_locator.press('Enter')
        except PlaywrightError as e:
            print(f' Failed to insert process number in the search field. Error: {type(e).__name__}')


    def search_text(self, target_name) -> list:
        # Select the iFrame
        try:
            viewer_frame = self.page.frame_locator("#ifrVisualizacao")
            content_frame = viewer_frame.frame_locator("iframe[id*='Html']")
            content_frame.locator("body").wait_for(state="visible", timeout=10000)
        except Exception as e:
            print(f'Could not find document iframe. terminating search')
            raise

        # Get the data about the process
        try:
            header_reference = content_frame.locator("p", has_text="Termo de Fomento").first
            first_p = header_reference.locator("xpath=following-sibling::p[1]")
            second_p = header_reference.locator("xpath=following-sibling::p[2]")

            data = {
                "line_1": first_p.inner_text().strip() if first_p.count() > 0 else None,
                "line_2": second_p.inner_text().strip() if second_p.count() > 0 else None
            }
            print(f"Paragraph 1: {data['line_1']}")
            print(f"Paragraph 2: {data['line_2']}")
            ### MAKE A FUNTION TO SEPARETE AND DELIVER THE DATA!

        except Exception as e:
            print(f'Could not find expected completion time. Error name: {type(e).__name__}\nError: {str(e)[:100]}')

        # Gets data about the receiving party
        try:
            dou_reference = content_frame.locator("p", has_text="Diário Oficial da União").first
            next_p_locator = dou_reference.locator("xpath=following-sibling::p[1]")

            if next_p_locator.count() > 0:
                target_text = next_p_locator.inner_text().strip()
                print(f"Texto encontrado: {target_text}")
            else:
                print("Nenhum parágrafo encontrado após a menção ao DOU.")
        except Exception as e:
            print(f'Could not find second paragraph text. Error name: {type(e).__name__}\nError: {str(e)[:100]}')

        # Get the object
        try:
            clausula_1 = content_frame.locator("p", has_text="CLÁUSULA PRIMEIRA ")
            vigencia_element = clausula_1.locator("xpath=following-sibling::p[1]").locator("strong, b").first

            objective_texto = vigencia_element.inner_text().strip()
            print(f"Objeto é: {objective_texto}")

        except Exception as e:
            print(f'Could not find objective. Error name: {type(e).__name__}\nError: {str(e)[:100]}')

        # Get the execution period
        try:
            clausula_3 = content_frame.locator("p", has_text="CLÁUSULA TERCEIRA")
            vigencia_element = clausula_3.locator("xpath=following-sibling::p[1]").locator("strong, b").first

            prazo_texto = vigencia_element.inner_text().strip()
            print(f"O prazo identificado é: {prazo_texto}")

        except Exception as e:
            print(f'Could not find expected completion time. Error name: {type(e).__name__}\nError: {str(e)[:100]}')

        # Get total value
        try:
            clausula_4 = content_frame.locator("p", has_text="CLÁUSULA QUARTA")
            valor_element = clausula_4.locator("xpath=following-sibling::p[1]").locator("strong, b")

            valor_total = valor_element.inner_text()
            print(f"O valor identificado é: {valor_total}")

        except Exception as e:
            print(f'Could not find transfer value. Error name: {type(e).__name__}\nError: {str(e)[:100]}')

        # Get reciving party data
        try:
            signature_div = content_frame.locator('div[unselectable="on"]').first
            paragraph_before = signature_div.locator("xpath=preceding-sibling::p[1]")

            if paragraph_before.count() > 0:
                text_content = paragraph_before.inner_text().strip()
                print(f"Text found before signature: {text_content}")
            else:
                print("Could not find a paragraph immediately before the signature block.")
        except Exception as e:
            print(f'Could not find transfer value. Error name: {type(e).__name__}\nError: {str(e)[:100]}')

        # Get signature date
        try:
            signature_p = content_frame.locator('div[unselectable="on"] p', has_text=target_name).first
            if signature_p.count() > 0:
                full_text = signature_p.inner_text()

                date_match = re.search(r"(\d{2}/\d{2}/\d{4})", full_text)
                if date_match:
                    extracted_date = date_match.group(1)
                    print(f"Found signature on {extracted_date} ")

        except Exception as e:
            print(f'Could not find transfer value. Error name: {type(e).__name__}\nError: {str(e)[:100]}')

        #sys.exit()
        return []

    def mark_as_done(self, df, numero_processo, phone, email):
        """Safely mark row as done in the DataFrame"""
        time.sleep(1)
        try:
            # Find the row index where the process number matches
            mask = df.iloc[:, 4] == numero_processo

            if mask.any():
                # Get the actual index position(s)
                idx_positions = df.index[mask]
                for idx in idx_positions:
                    df.at[idx, 'Telefone'] = phone
                    df.at[idx, 'Email'] = email

                print(f"✅ Updated process {numero_processo}: Phone={phone}, Email={email}")
                return True
            else:
                print(f"❌ Process number {numero_processo} not found in DataFrame")
                return False

        except Exception as err:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Error in mark_as_done:{type(err).__name__}\n{str(err)[:100]}\n")
            return False


    def loop_de_pesquisa(self, df, numero_processo: str):
        print(f'🔍 Starting data extraction loop'.center(50, '-'), '\n')
        xpath = [
            '//*[@id="menu_link_330887298_1035188630"]',
            '//*[@id="menu_link_330887298_-442619135"]',
            '//*[@id="grupo_abas"]/a[3]',
            '//*[@id="menu_link_2144784112_100344749"]',
            '//*[@id="menu_link_2144784112_2061164"]',
            '//*[@id="dados-cauc"]/siconv-historico/div/div[4]/button[2]',
            '//*[@id="breadcrumbs"]/a[2]'
        ]
        try:
            self.campo_pesquisa(numero_processo=numero_processo)
            for x in xpath:
                time.sleep(2.5)
                self.page.locator(x).click(timeout=10000)

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
            raise


    def search_in_sei_tree(self):
        search_text = ['Termo de Fomento', 'Convênio']
        try:
            tree_frame = self.page.frame_locator("#ifrArvore")
            # Expand all fields
            tree_frame.get_by_title("Abrir todas as Pastas").click()

            pattern = re.compile(r"^(Termo de Fomento|Convênio)", re.IGNORECASE)

            target_link = tree_frame.locator("span, a").filter(has_text=pattern).first
            target_link.wait_for(state="attached", timeout=5000)
            target_link.click()

            print("Document link clicked successfully.")
        except:
            print(f"Could not find '{search_text}' in the tree.")

        return None

    @staticmethod
    def save_to_excel(df, caminho_arquivo_fonte, sheet_name='Sheet1'):
        try:
            # Create backup of original file
            backup_path = caminho_arquivo_fonte.replace('.xlsx', '_backup.xlsx')
            if os.path.exists(caminho_arquivo_fonte):
                import shutil
                shutil.copy2(caminho_arquivo_fonte, backup_path)
                print(f"✅ Backup created: {backup_path}")

            # Save using 'w' mode to overwrite the entire file
            with ExcelWriter(caminho_arquivo_fonte, engine='openpyxl', mode='w') as writer:
                df.to_excel(writer, index=False, sheet_name=sheet_name)

            print(f"✅ Successfully saved {len(df)} rows to Excel.")
            return True

        except PermissionError:
            print(f"❌ Permission denied: Please close the Excel file '{caminho_arquivo_fonte}' and try again.")
            return False
        except Exception as erro:
            exc_type, exc_value, exc_tb = sys.exc_info()
            line_number = exc_tb.tb_lineno
            if exc_tb.tb_next:
                line_number = exc_tb.tb_next.tb_lineno
            print(f"Error occurred at line: {line_number}")
            print(f'❌ Failed to save Excel file. Error:{type(erro).__name__}\n{str(erro)}')
            return False

    @staticmethod
    def extrair_dados_excel(caminho_arquivo_fonte):
        try:
            complete_data_frame = pd.read_excel(caminho_arquivo_fonte, dtype=str, sheet_name=None)
            sheet_names_list = list(complete_data_frame.keys())

            print("✅ Sheet Names Found:")
            for name in sheet_names_list:
                print(f"- {name}")

            data_frame = complete_data_frame['Sheet1']
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

        pattern = r'^\d{5}/\d{4}'

        if '_' in numero_proposta:
            numero_proposta_fixed = numero_proposta.replace('_', '/')
            return numero_proposta_fixed

        if re.findall(pattern, numero_proposta):
            return numero_proposta
        else:
            return False


def main():
    robo = PWRobo()

    robo.search(numero_processo='71000.117864/2025-74')

    robo.search_in_sei_tree()

    robo.search_text('PAULO HENRIQUE PERNA CORDEIRO')



if __name__ == '__main__':
    main()















