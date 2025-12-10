import os.path
import sys
import re

from pandas import ExcelWriter
import pandas as pd

from colorama import Fore

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError
from playwright.sync_api import sync_playwright


class BreakInnerLoop(Exception):
    pass


class PWRobo:
    def __init__(self, cdp_url: str = "http://localhost:9222"):
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.connect_over_cdp(cdp_url)
        self.context = self.browser.contexts[0] if self.browser.contexts else self.browser.new_context()
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()

        # Enable resource blocking for faster performance
        self.block_rss()

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
        """Navigates through the system tabs to the process search page."""
        try:
            self.page.locator('xpath=//*[@id="logo"]/a').click(timeout=800)
        except PlaywrightTimeoutError:
            print('Already on the initial page of transferegov discricionárias.')
        except PlaywrightError as e:
            print(Fore.RED + f'🔄❌ Failed to reset.\nError: {type(e).__name__}\n{str(e)[:100]}')

        xpaths = ['xpath=//*[@id="menuPrincipal"]/div[1]/div[3]', 'xpath=//*[@id="contentMenu"]/div[1]/ul/li[2]/a']
        try:
            for xpath in xpaths:
                self.page.locator(xpath).click(timeout=10000)
        except PlaywrightError as e:
            print(Fore.RED + f'🔴📄 Instrument unavailable. \nError: {e}')
            sys.exit(1)

    def campo_pesquisa(self, numero_processo):
        try:
            campo_pesquisa_locator = self.page.locator('xpath=//*[@id="consultarNumeroProposta"]')
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

    def busca_endereco(self):
        self.page.wait_for_timeout(500)
        try:
            print(f'🔍 Starting data search'.center(50, '-'), '\n')

            # Click the detail button
            self.page.locator('input#form_submit[value="Detalhar"]').click()

            # Wait for the page to load
            self.page.wait_for_timeout(2000)

            # Initialize with empty values
            phone = ""
            email = ""

            # Check if elements exist before trying to extract text
            phone_locator = self.page.locator('#txtTelefone')
            email_locator = self.page.locator('#txtEmail')

            phone_exists = phone_locator.count() > 0
            email_exists = email_locator.count() > 0

            print(f"🔍 Element check - Phone field exists: {phone_exists}, Email field exists: {email_exists}")

            # Extract phone if available
            if phone_exists:
                try:
                    phone_locator.wait_for(timeout=3000)
                    phone = phone_locator.inner_text(timeout=2000).strip()
                    print(f"📞 Phone found: '{phone}'")
                except PlaywrightTimeoutError:
                    print("⚠️ Phone element exists but content not loaded or empty")
                except Exception as e:
                    print(f"⚠️ Error extracting phone: {e}")
            else:
                print("📞 Phone field not found on page")

            # Extract email if available
            if email_exists:
                try:
                    email_locator.wait_for(timeout=3000)
                    email = email_locator.inner_text(timeout=2000).strip()
                    print(f"📧 Email found: '{email}'")
                except PlaywrightTimeoutError:
                    print("⚠️ Email element exists but content not loaded or empty")
                except Exception as e:
                    print(f"⚠️ Error extracting email: {e}")
            else:
                print("📧 Email field not found on page")

            # Check if we got any data
            if not phone and not email:
                print("ℹ️ No contact information found on this page")
            elif phone and not email:
                print("ℹ️ Only phone number found")
            elif email and not phone:
                print("ℹ️ Only email found")
            else:
                print("✅ Both phone and email found")

            # Back button
            try:
                self.page.locator('xpath=//*[@id="lnkConsultaAnterior"]').click(timeout=8000)
            except:
                print("⚠️ Could not click back button, navigating manually")
                # Alternative navigation if back button fails
                self.page.go_back()

            return phone, email

        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Unexpected error in busca_endereco: {type(e).__name__} - {str(e)[:100]}")
            # Ensure we navigate back even on error
            self.page.locator('xpath=//*[@id="lnkConsultaAnterior"]').click(timeout=2000)
            return "", ""


    def mark_as_done(self, df, numero_processo, phone, email):
        """Safely mark row as done in the DataFrame"""
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

        try:
            self.campo_pesquisa(numero_processo=numero_processo)
            phone, email = self.busca_endereco()

            if phone is not None and email is not None:
                success = self.mark_as_done(df=df, numero_processo=numero_processo, phone=phone, email=email)
                if success:
                    self.page.locator('xpath=//*[@id="breadcrumbs"]/a[2]').click(timeout=3000)
                    return True
            return False

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

    # --- Fixed Save Function ---
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

        if  re.findall(pattern, numero_proposta):
            return numero_proposta
        else:
            return False



def main() -> None:
    dir_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\fabi_DFP\Convênios - Pendências Celebração (24.11) - Copia.xlsx')

    try:
        robo = PWRobo()
    except Exception as e:
        print(f"\n‼️ Fatal error starting the robot: {e}")
        sys.exit("Stopping the program.")

    # Load DataFrame
    df = robo.extrair_dados_excel(caminho_arquivo_fonte=dir_path)
    if df is None:
        print("❌ Failed to load Excel file. Exiting.")
        return

    # Make sure the required columns exist
    if 'Telefone' not in df.columns:
        df['Telefone'] = ''
    if 'Email' not in df.columns:
        df['Email'] = ''

    robo.consulta_proposta()

    successful_updates = 0
    skipped_count = 0

    for idx, row in df.iterrows():
        numero_processo_temp = row.iloc[4]
        numero_processo = robo.fix_prop_num(numero_processo_temp)

        if not numero_processo:
            continue

        # Get current phone and email values
        current_phone = str(row['Telefone']) if pd.notna(row['Telefone']) else ""
        current_email = str(row['Email']) if pd.notna(row['Email']) else ""

        # Skip if both phone and email already have data
        if current_phone or current_email:
            print(f"⏭️  Skipping {numero_processo} - already processed (both phone and email exist)")
            skipped_count += 1
            continue

        print(f"\n{'⚡' * 3}🚀 EXECUTING PROPOSAL: {numero_processo} 🚀{'⚡' * 3}".center(70, '='), '\n')
        print(f"Current progress {idx/len(df):.2%}.")

        try:
            success = robo.loop_de_pesquisa(df=df, numero_processo=numero_processo)
            if success:
                successful_updates += 1

            # Save every 10 rows or at the end
            if (idx + 1) % 10 == 0 or idx == len(df) - 1:
                print(f"💾 Saving progress... ({idx + 1}/{len(df)} rows processed)")
                if robo.save_to_excel(df=df, caminho_arquivo_fonte=dir_path):
                    print(f"✅ Successfully saved {successful_updates} updates so far, {skipped_count} skipped")

        except BreakInnerLoop:
            print("⚠️ Stopping this unique_values loop early.")
            # Save progress before breaking
            robo.save_to_excel(df=df, caminho_arquivo_fonte=dir_path)
            break
        except KeyboardInterrupt:
            print("Script stopped by user (Ctrl+C). Saving progress...")
            robo.save_to_excel(df=df, caminho_arquivo_fonte=dir_path)
            sys.exit(0)
        except Exception as e:
            exc_type, exc_value, exc_tb = sys.exc_info()
            print(f"Error occurred at line: {exc_tb.tb_lineno}")
            print(f"❌ Failed to execute script. Error: {type(e).__name__}\n{str(e)}")
            # Save progress on error
            robo.save_to_excel(df=df, caminho_arquivo_fonte=dir_path)

    print(f"\n🎉 Processing complete!")
    print(f"📊 Summary: {successful_updates} updated, {skipped_count} skipped, {len(df)} total")

if __name__ == "__main__":
    main()
