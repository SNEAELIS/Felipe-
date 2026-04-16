import os
import sys
import logging.handlers
import logging
import time

import pandas as pd

from datetime import datetime

from fontTools.misc.cython import returns
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError
from playwright.sync_api import sync_playwright


class PWRobo:
    def __init__(self, webinar_invite: str = None, text_tf: str = None, text_conv: str = None,
                 cdp_url: str = "http://localhost:9222"):
        # Standard text for term of promotion proposal
        self.feedback_txt_one = text_tf

        # Standard text for agreement proposal
        self.feedback_txt_two = text_conv

        # Connect to existing browser via Chrome DevTools Protocol (CDP)
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.connect_over_cdp(cdp_url)

        # Get all browser contexts (browser windows/profiles)
        self.context = self.browser.contexts[0] if self.browser.contexts else self.browser.new_context()

        # Get all open pages (tabs)
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()

        # Defines Logger
        self.logger = self.setup_logger()

        valid_tabs = self.skip_chrome_tab_search()

        # Switch to a specific tab if needed
        if valid_tabs:
            self.page = valid_tabs[0]  # Switch to first valid tab
            self.page.bring_to_front()  # Focus the tab

        print(f"✅ Connected to existing Chrome instance via Playwright. Connected to page: {self.page.url}")


    def skip_chrome_tab_search(self):
        """
        Skips Chrome-specific tabs (like chrome:// URLs) and returns only valid pages.

        Returns:
            list: List of valid page objects (tabs) that are not Chrome internal pages
        """
        # Get all pages in the current context
        all_pages = self.context.pages

        # Filter out Chrome internal pages
        valid_pages = []
        for page in all_pages:
            url = page.url
            # Skip Chrome internal pages (chrome://, chrome-extension://, etc.)
            if not (url.startswith("chrome://") or url.startswith("chrome-extension://")):
                valid_pages.append(page)

        # Optional: Switch to the first valid page
        if valid_pages:
            self.page = valid_pages[0]
            print(f"✅ Switched to valid page: {self.page.url}")

        return valid_pages


    @staticmethod
    def setup_logger(level=logging.INFO):
        log_file_path = r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\Log_pareceres'

        # Create logs directory if it doesn't exist
        log_file_name = f'log_Pareceres_email{datetime.now().strftime('%d_%m_%Y')}.log'

        # Sends to specific directory
        log_file = os.path.join(log_file_path, log_file_name)

        if not os.path.exists(log_file_path):
            os.makedirs(log_file_path)
            print(f"✅ Directory created/verified: {log_file_path}")

        logger = logging.getLogger()
        if logger.handlers:
            return logger

        # Avoid adding handlers multiple times
        logger.setLevel(level)

        formatter = logging.Formatter(
            '%(asctime)s | | %(message)s\n' + '─' * 100,
            datefmt='%Y-%m-%d  %H:%M'
        )

        # File handler with rotation
        file_handler = logging.handlers.RotatingFileHandler(
            log_file,
            maxBytes=10485760,
            backupCount=5,
            encoding='utf-8'
        )
        file_handler.setFormatter(formatter)

        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)

        if os.path.exists(log_file):
            print(f"🎉 SUCCESS! Log file created at: {log_file}")
            print(f"📊 File size: {os.path.getsize(log_file)} bytes")
            return logger
        else:
            print(f"❌ File not created at: {log_file}")


    @staticmethod
    def mark_proposal_done(df: pd.DataFrame, proposal_num: str, file_path: str, sheet_name: str) -> bool:
        """
        Marks a specific proposal number as 'Done' in a 'Status' column at the row index where the proposal number matches.

        Args:
            df (pd.DataFrame): The original DataFrame.
            proposal_num (str): The specific proposal number to mark.
            file_path (str): Path to save the Excel file.
            sheet_name (str): Name of the sheet to save.

        Returns:
            bool: True if successful, False if proposal_num not found.
        """

        # Find the row index where the proposal number matches
        # Assuming the proposal number is in a column named 'Proposal_Num' or similar
        # Adjust column name as needed
        mask = df["Proposta"] == proposal_num  # Change 'Proposal_Num' to your actual column name

        if not mask.any():
            return False  # Proposal number not found

        # Mark the matching row(s) as 'Feito'
        df.loc[mask, "Feito"] = "Feito"

        # Save the DataFrame to Excel
        df.to_excel(file_path, sheet_name=sheet_name, index=False)

        return True


    def block_rss(self):
        def route_handler(route, request):
            if request.resource_type in ["image", "stylesheet", "font"]:
                route.abort()
            else:
                route.continue_()
        self.page.route('**/*', route_handler)


    def init_search(self):
        try:
            # In the landing page access proposal sub menu
            self.page.click("xpath=//*[@id='menuPrincipal']/div[1]/div[3]", timeout=15000)
            # Search for proposal
            self.page.click("xpath=//div[@id='contentMenu']//a[normalize-space()='Consultar Propostas']",
                            timeout=15000)

        except PlaywrightTimeoutError as te:
            print(f"❗⏱️ Timeout occurred during initial search: {str(te)[:100]}\nErro name"
                  f":{type(te).__name__}")
        except PlaywrightError as pe:
            print(f"❗🧩 Playwright-specific error: {str(pe)[:100]}\nErro name:{type(pe).__name__}")
        except Exception as e:
            print(f"🚨🚨 An unexpected error occurred while initiating search: {str(e)[:100]}\nErro name"
                  f":{type(e).__name__}")


    def loop_search(self, prop_num: str, idx: int):
        try:
            self.logger.info(f"Pesquisando índice:{idx}, proposta: {prop_num} ")
            # Select desired proposal
            self.page.fill("#consultarNumeroProposta", f"{prop_num}")  # number input
            self.page.click("xpath=(//input[@id='form_submit'])[1]", timeout=15000)  # click to consult
            self.page.click("div[class='numeroProposta'] a", timeout=15000)  # click to select

            # Navigate to Action Plan
            self.page.click("div[id='div_997366806'] span span", timeout=15000)
            # Navigate to Feedback
            self.page.click("a[id='menu_link_997366806_-231259270'] div[class='inactiveTab'] span span",
                            timeout=15000)

        except PlaywrightTimeoutError as te:
            print(f"❗⏱️ Timeout occurred during search loop: {str(te)[:100]}\nErro name:{type(te).__name__}")
        except PlaywrightError as pe:
            print(f"❗🧩 Playwright-specific error: {str(pe)[:100]}\nErro name:{type(pe).__name__}")
        except Exception as e:
            print(f"🚨🚨 An unexpected error occurred under loop search: {str(e)[:100]}\nErro name"
                  f":{type(e).__name__}")


    def insert_feedback(self, type_txt: int):
        #return True
        try:
            # All form submit in the page
            inserir_btn_all = self.page.locator("xpath=//input[@id='form_submit']").all()
            inserir_btn_list = [x for x in inserir_btn_all if x.is_visible(timeout=15000)]
            print(f"Debugg print for button list: \n"
                  f"{[str(s)[:20] for s in inserir_btn_list]}\n")
            elements_to_use = []
            found_elements = False

            for i, element in enumerate(inserir_btn_list):
                try:
                    element_value = element.get_attribute("value",timeout=1500)
                    if element_value in ['Inserir Parecer da Proposta']:
                        elements_to_use.append(element)
                        found_elements = True
                        break
                    else:
                        continue
                except Exception:
                    if len(elements_to_use) == 0:
                        print(f"\nNo element found. Trying to initiate analisys\n")
                        self.open_req(type_txt=type_txt)
                        return False
                    continue

            if not found_elements or len(elements_to_use) == 0:
                print("\nNo suitable elements found after scanning all buttons")
                return False

            elements_to_use[-1].click(timeout=15000)

            txt_field = self.page.locator("xpath=//*[@id='emitirParecerParecer']")
            if type_txt == 0:
                print('Colocando texto de termo de fomento')
                # Text for term of promotion proposal
                txt_field.fill(f"{self.feedback_txt_one}", timeout=15000)
                # Press "Issue Feedback" button
                try:
                    self.page.click('input[value="Emitir Parecer"]', timeout=15000)

                    complete_ = self.page.wait_for_selector("xpath=//div[@class='message']",
                                                            timeout=15000)
                    if complete_:
                        print("✅🤝 Parecer inserido com sucesso. Convite Webinar")
                        self.logger.info('Parecer inserido com sucesso. Convite Webinar')
                        # Press return button
                        self.page.click("xpath=//input[@id='form_submit']", timeout=15000)  # click to consult
                        return True


                except PlaywrightTimeoutError:
                    print('Wrong element trying second element')
                    try:
                        self.page.click(r"xpath=(//input[@id='form_submit_emitir'])[1]",
                                        timeout=15000)
                        complete_ = self.page.wait_for_selector("xpath=//div[@class='message']",
                                                                timeout=15000)
                        if complete_:
                            print("✅🤝 Parecer inserido com sucesso")
                            self.logger.info('Parecer inserido com sucesso. Convite Webinar')
                            # Press return button
                            self.page.click("xpath=//input[@id='form_submit']",
                                            timeout=15000)  # click to consult
                            return True

                    except PlaywrightTimeoutError:
                        self.logger.info('Erro ao inserir parecer')
                        sys.exit(1)
            else:
                # Text for agreement proposal
                txt_field.fill(f"{self.feedback_txt_two}", timeout=15000)
                # Press "Issue Feedback" button
                try:
                    self.page.click("body > div:nth-child(16) > div:nth-child(18) > div:nth-child(9) > "
                                    "div:nth-child(1) > div:nth-child(1) > form:nth-child(1) > table:nth-child("
                                    "6) > tbody:nth-child(1) > tr:nth-child(5) > td:nth-child(2) > "
                                    "input:nth-child(1)",
                                    timeout=15000)
                    complete_ = self.page.wait_for_selector("xpath=//div[@class='message']",
                                                            timeout=15000)
                    if complete_:
                        print("✅🤝 Parecer inserido com sucesso")
                        self.logger.info('Parecer inserido com sucesso')
                        # Press return button
                        self.page.click("xpath=//input[@id='form_submit']",
                                        timeout=15000)  # click to consult
                        return True

                except PlaywrightTimeoutError:
                    print('Wrong element trying second element')
                    try:
                        self.page.click(r"xpath=(//input[@id='form_submit_emitir'])[1]",
                                        timeout=15000)
                        complete_ = self.page.wait_for_selector("xpath=//div[@class='message']",
                                                                timeout=15000)
                        if complete_:
                            print("✅🤝 Parecer inserido com sucesso")
                            self.logger.info('Parecer inserido com sucesso')
                            # Press return button
                            self.page.click("xpath=//input[@id='form_submit']",
                                            timeout=15000)  # click to consult
                            return True

                    except PlaywrightTimeoutError:
                        self.logger.info('Erro ao inserir parecer')
                        sys.exit(1)

            return False


        except PlaywrightTimeoutError as te:
            print(f"❗⏱️ Timeout occurred during feedback insertion: {str(te)[:100]}\nErro name"
                  f":{type(te).__name__}")
            self.logger.info('Erro ao inserir parecer')
            return False
        except PlaywrightError as pe:
            print(f"❗🧩 Playwright-specific error: {str(pe)[:100]}\nErro name:{type(pe).__name__}")
            self.logger.info('Erro ao inserir parecer')
            return False
        except Exception as e:
            print(f"🚨🚨 An unexpected error occurred while inserting feedback: {str(e)[:100]}\nErro name"
                  f":{type(e).__name__}")
            self.logger.info('Erro ao inserir parecer')
            return False


    def open_req(self, type_txt: int):
        try:
            # Go to Proposal Data page
            self.page.click("xpath=//span[contains(text(),'Dados da Proposta')]")
            # Open Data tab
            self.page.click("xpath=//div[@class='inactiveTab']//span//span[contains(text(),'Dados')]")
            # Click Initiate Analysis
            anls_init = self.page.wait_for_selector("xpath=//tbody//tr//input[2]", timeout=15000)
            if anls_init == "Analisar Plano de Trabalho":
                return
            anls_init.click()
            anls_confirm = self.page.wait_for_selector("xpath=//tbody//tr//input[1]", timeout=15000)
            anls_confirm.click()

            anls_sccs = self.page.wait_for_selector('xpath=//*[@id="messages"]/div/div',
                                                    timeout=500)
            if anls_sccs:
                print("Successfully initiated analysis")
                self.logger.info("Abertura de análise concluída com sucesso")

                # Navigate to Action Plan
                self.page.click("div[id='div_997366806'] span span", timeout=15000)
                # Navigate to Feedback
                self.page.click("a[id='menu_link_997366806_-231259270'] div[class='inactiveTab'] span span",
                                timeout=15000)
                self.insert_feedback(type_txt=type_txt)
            else:
                sys.exit()

        except Exception as e:
            print(f"🚨🚨 An unexpected error occurred while trying to initiate analysis: {str(e)[:100]}\nErro "
                  f"name"
                  f":{type(e).__name__}")
            self.logger.info('Erro ao iniciar análise')

            sys.exit()


    def reset(self):
        try:
            self.page.click("xpath=//*[@id='breadcrumbs']/a[2]", timeout=150000)
            self.page.wait_for_selector("xpath=//div[normalize-space()='Propostas']")
        except PlaywrightTimeoutError as te:
            print(f"❗⏱️ Timeout occurred during reset: {str(te)[:100]}\nErro name:{type(te).__name__}")
        except PlaywrightError as pe:
            print(f"❗🧩 Playwright-specific error: {str(pe)[:100]}\nErro name:{type(pe).__name__}")
        except Exception as e:
            print(f"🚨🚨 An unexpected error occurred while reseting: {str(e)[:100]}\nErro name"
                  f":{type(e).__name__}")


    def land_page(self):
        try:
            print("Trying to reset")
            icon = self.page.wait_for_selector("xpath=//*[@id='logo']/a/img", timeout=150000)
            if icon.is_enabled():
                icon.click()
                time.sleep(1.5)

        except PlaywrightTimeoutError as te:
            print(f"❗⏱️ Timeout occurred at land page link location: {str(te)[:100]}\nErro name"
                  f":{type(te).__name__}")
            return False
        except PlaywrightError as pe:
            print(f"❗🧩 Playwright-specific error: {str(pe)[:100]}\nErro name:{type(pe).__name__}")
            return False
        except Exception as e:
            print(f"🚨🚨 An unexpected error occurred at land_page: {str(e)[:100]}\nErro name"
                  f":{type(e).__name__}")
            return False


    def requirements(self):
        threshhold_date = datetime(2025, 12, 18)

        txt_ = (f"Para inserção das documentações pendentes, que estão relacionadas no parecer emitido em "
                f"08/10/2025 (aba pareceres), se for o caso.")

        # Select Requisites
        self.page.click("xpath=//div[@id='div_2144784112']//span//span[contains(text(),'Requisitos')]"
                        , timeout=15000)
        # SelectRequisites for celebration
        self.page.click("xpath=//span[contains(text(),'Requisitos para Celebração')]",
                        timeout=15000)

        # Select Historic table
        self.page.wait_for_selector("xpath=//div[@id='divFormulario']//div[5]", timeout=15000)

        table_ = self.page.locator("xpath=//div[@id='divFormulario']//div[5]")
        rows = table_.locator('tr')

        for i in range(1, rows.count()):
            row = rows.nth(i)

            cells = row.locator('td, th')
            # Only proceed if we have at least 2 cells
            if cells.count() >= 2:
                event_ = cells.nth(0).text_content().strip()
                date_txt = cells.nth(2).text_content().strip().split()[0]
                date_obj = datetime.strptime(date_txt, "%d/%m/%Y")

                if event_ == "Complementação Solicitada" and date_obj >= threshhold_date:
                    print('Já existe solicitação de complementação em data posterior à 08/10/2025')
                    self.logger.info('Já existe solicitação de complementação em data posterior à 08/10/2025')
                    return

        # Request complementation button
        self.page.click("xpath=//input[@id='formRequisitosDocumentos:_idJsp159']", timeout=15000)
        #
        txt_field = "xpath=//textarea[@id='formSolicitacaComplementacao:observacao']"
        self.page.fill(selector=txt_field, value=f"{txt_}", timeout=15000)

        self.page.click("xpath=//input[@id='formSolicitacaComplementacao:_idJsp23']")

        complete_ = self.page.wait_for_selector("xpath=//div[@class='messages']//div", timeout=15000)
        if complete_:
            print("✅🤝 Solicitação de complementação enviada com sucesso")
            self.logger.info('Solicitação de complementação enviada com sucesso')

    def error_message(self):
        try:
            self.page.wait_for_selector('//*[@id="popUpLayer2"]', timeout=500)
            close_button = self.page.locator('#popUpLayer2 img[src*="close.gif"]')
            if close_button:
                close_button.wait_for(state="visible", timeout=1000)
                close_button.click()
            return True

        except PlaywrightTimeoutError:
            print(f"Mensagem de erro não encontrada")
            return False
        except Exception as e:
            print(f"🚨🚨 An unexpected error occurred with error message: {str(e)[:100]}\nErro name"
                  f":{type(e).__name__}")
            return False

    @staticmethod
    def create_pending_list(pending_text) -> list:
        return []


# convenio
def main():
    xlsx_source_path = r"C:\Users\felipe.rsouza\Downloads\sneaelis_tabela_202604011316 - Copia.xlsx"

    # Texts to be parsed

    text_tf = rf''' 
📢 Retificação de Notificação

Pedimos desculpas pelo envio indevido de uma mensagem de notificação anteriormente encaminhada.

Informamos que, após verificação, constatamos que esta entidade não precisa realizar o cadastro, não sendo necessária qualquer ação em relação à mensagem recebida.

Esclarecemos que a mensagem anterior era destinada apenas às entidades que possuem Convênios.
                            '''

    # Get excel file
    excel_file = pd.ExcelFile(xlsx_source_path)
    # Get all sheets inside the file and store in a list
    sheet_names = excel_file.sheet_names

    # Initiate automation instance
    robo = PWRobo(text_tf=text_tf)

    robo.land_page()
    robo.init_search()

    for i, sheet in enumerate(sheet_names):
        no_do_ = ['Completa', 'Planilha', 'Termo de Fomento']
        if sheet in no_do_:
            continue
        print(f"\n{'<' * 7}📄 Loading sheet [{i}/{len(sheet_names)}]: '{sheet}'{'>' * 7}"
              .center(80, '-'))
        try:
            # Create the DataFrame with source data
            df = pd.read_excel(xlsx_source_path, dtype=str, sheet_name=sheet)
            print(f"\n✅ Sheet '{sheet}' loaded successfully with {len(df)} rows.\n")

            for idx, row in df.iterrows():
                try:
                    proposal_done = row["Feito"]

                    prop_num = row['Proposta']
                    if proposal_done == "Feito":
                        print(f"Proposal: {prop_num} already filled")
                        continue
                    if pd.isna(prop_num):
                        continue
                    # pending_req = row['XXXXX']
                    #
                    # tech_phone_number = row['XXXXX']
                    #
                    # tech_email = row['XXXXX']
                    #
                    # pending_list = robo.create_pending_list(pending_req)


                    print("\n", f"{'⚡' * 3}🚀 EXECUTING PROPOSAL: {prop_num}, index: {idx} "
                               f"🚀{'⚡' * 3}".center(70,
                           '='), '\n')

                    robo.loop_search(prop_num, idx)
                    feedback_ = robo.insert_feedback(type_txt=i) # REFACTOR THE RETURN VALUES!!!!

                    if feedback_:
                        robo.mark_proposal_done(df=df, proposal_num=prop_num, file_path=xlsx_source_path,
                                                   sheet_name=sheet)
                        #robo.requirements()
                    robo.reset()
                    try:
                        if robo.error_message():
                            robo.land_page()
                            robo.init_search()
                    except Exception as e:
                        print(
                           f"🚨🚨 An unexpected error occurred during error handdling in main:"
                           f" {str(e)[:100]}\nErro name"
                           f":{type(e).__name__}")

                except Exception as e:
                    print(f"🚨🚨 An unexpected error occurred during main: {str(e)[:100]}\nErro name"
                          f":{type(e).__name__}")
        except Exception as e:
            print(f"🚨🚨 Failed to load or process sheet '{sheet}': {str(e)[:100]} (Error: {type(e).__name__})")


if __name__ == "__main__":
    main()
