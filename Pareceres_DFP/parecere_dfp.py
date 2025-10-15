import os
import sys
import logging.handlers
import logging
import time

import pandas as pd

from datetime import datetime

from playwright.sync_api import TimeoutError as PlaywrightTimeoutError, Error as PlaywrightError
from playwright.sync_api import sync_playwright


class PWRobo:
    def __init__(self, text_tf:str, text_conv:str, cdp_url:str ="http://localhost:9222"):
        # Standard text for term of promotion proposal
        self.feedback_txt_one = text_tf
        # Standard text for agreement proposal
        self.feedback_txt_two = text_conv

        self.playwright = sync_playwright().start()
        # Connect to existing browser via Chrome DevTools Protocol (CDP)
        self.browser = self.playwright.chromium.connect_over_cdp(cdp_url)
        # Get all browser contexts (browser windows/profiles)
        self.context = self.browser.contexts[0] if self.browser.contexts else self.browser.new_context()
        # Get all open pages (tabs)
        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        #self.block_rss()

        # Defines Logger
        self.logger = self.setup_logger()

        print(f"✅ Connected to existing Chrome instance via Playwright. Connected to page: {self.page.url}")


    def setup_logger(self, level=logging.INFO):
        log_file_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                         r'Social\Teste001\fabi_DFP')

        # Create logs directory if it doesn't exist
        log_file_name = f'log_Pareceres_{datetime.now().strftime('%d_%m_%Y')}.log'

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
            self.page.click("xpath=//*[@id='menuPrincipal']/div[1]/div[3]", timeout=5000)
            # Search for proposal
            self.page.click("xpath=//div[@id='contentMenu']//a[normalize-space()='Consultar Propostas']",
                            timeout=5000)

        except PlaywrightTimeoutError as te:
            print(f"❗⏱️ Timeout occurred during initial search: {str(te)[:100]}\nErro name"
                  f":{type(te).__name__}")
        except PlaywrightError as pe:
            print(f"❗🧩 Playwright-specific error: {str(pe)[:100]}\nErro name:{type(pe).__name__}")
        except Exception as e:
            print(f"🚨🚨 An unexpected error occurred: {str(e)[:100]}\nErro name:{type(e).__name__}")


    def loop_search(self, prop_num: str):
        try:
            # Select desired proposal
            self.page.fill("#consultarNumeroProposta", f"{prop_num}")  # number input
            self.page.click("xpath=(//input[@id='form_submit'])[1]", timeout=5000)  # click to consult
            self.page.click("div[class='numeroProposta'] a", timeout=5000)  # click to select

            # Navigate to Action Plan
            self.page.click("div[id='div_997366806'] span span", timeout=5000)
            # Navigate to Feedback
            self.page.click("a[id='menu_link_997366806_-231259270'] div[class='inactiveTab'] span span",
                            timeout=5000)

        except PlaywrightTimeoutError as te:
            print(f"❗⏱️ Timeout occurred during search loop: {str(te)[:100]}\nErro name:{type(te).__name__}")
        except PlaywrightError as pe:
            print(f"❗🧩 Playwright-specific error: {str(pe)[:100]}\nErro name:{type(pe).__name__}")
        except Exception as e:
            print(f"🚨🚨 An unexpected error occurred: {str(e)[:100]}\nErro name:{type(e).__name__}")


    def insert_feedback(self, type_txt: int):
        try:
            # All form submit in the page
            inserir_btn_all = self.page.locator("xpath=//input[@id='form_submit']").all()
            inserir_btn_list = [x for x in inserir_btn_all if x.is_visible(timeout=5000)]
            elements_to_use = []

            for i, element in enumerate(inserir_btn_list):
                try:
                    element_value = element.get_attribute("value", timeout=2000)
                    print(f'{element_value}')
                    if element_value in ['Inserir Parecer da Proposta']:
                        elements_to_use.append(element)
                    else:
                        continue
                except Exception as e:
                    print(f"Element {i}: Error inspecting - {e}")
                    continue

            elements_to_use[-1].click(timeout=5000)

            txt_field = self.page.locator("xpath=//*[@id='emitirParecerParecer']")
            if type_txt == 1:
                print('Colocando texto de termo de fomento')
                # Text for term of promotion proposal
                txt_field.fill(f"{self.feedback_txt_one}", timeout=1000)
                # Press "Issue Feedback" button
                try:
                    self.page.click("body > div:nth-child(16) > div:nth-child(18) > div:nth-child(9) > "
                                    "div:nth-child(1) > div:nth-child(1) > form:nth-child(1) > table:nth-child("
                                    "6) > tbody:nth-child(1) > tr:nth-child(5) > td:nth-child(2) > "
                                    "input:nth-child(1)",
                                    timeout=5000)

                    complete_ = self.page.wait_for_selector("xpath=//div[@class='message']",
                                                            timeout=5000)
                    if complete_:
                        print("✅🤝 Parecer inserido com sucesso")
                        self.logger.info('Parecer inserido com sucesso')

                except PlaywrightTimeoutError:
                    print('Wrong element trying second element')
                    try:
                        self.page.click(r"xpath=(//input[@id='form_submit_emitir'])[1]",
                                        timeout=1000)
                        complete_ = self.page.wait_for_selector("xpath=//div[@class='message']",
                                                                timeout=10000)
                        if complete_:
                            print("✅🤝 Parecer inserido com sucesso")

                    except PlaywrightTimeoutError:
                        self.logger.info('Erro ao inserir parecer')
                        sys.exit(1)
            else:
                # Text for agreement proposal
                txt_field.fill(f"{self.feedback_txt_two}", timeout=1000)
                # Press "Issue Feedback" button
                try:
                    self.page.click("body > div:nth-child(16) > div:nth-child(18) > div:nth-child(9) > "
                                    "div:nth-child(1) > div:nth-child(1) > form:nth-child(1) > table:nth-child("
                                    "6) > tbody:nth-child(1) > tr:nth-child(5) > td:nth-child(2) > "
                                    "input:nth-child(1)",
                                    timeout=5000)
                    complete_ = self.page.wait_for_selector("xpath=//div[@class='message']",
                                                            timeout=5000)
                    if complete_:
                        print("✅🤝 Parecer inserido com sucesso")
                        self.logger.info('Parecer inserido com sucesso')

                except PlaywrightTimeoutError:
                    print('Wrong element trying second element')
                    try:
                        self.page.click(r"xpath=(//input[@id='form_submit_emitir'])[1]",
                                        timeout=1000)
                        complete_ = self.page.wait_for_selector("xpath=//div[@class='message']",
                                                                timeout=5000)
                        if complete_:
                            print("✅🤝 Parecer inserido com sucesso")
                            self.logger.info('Parecer inserido com sucesso')

                    except PlaywrightTimeoutError:
                        self.logger.info('Erro ao inserir parecer')
                        sys.exit(1)

            # Press return button
            self.page.click("xpath=//input[@id='form_submit']", timeout=5000)  # click to consult
            return True

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
            print(f"🚨🚨 An unexpected error occurred whie inserting feedback: {str(e)[:100]}\nErro name"
                  f":{type(e).__name__}")
            self.logger.info('Erro ao inserir parecer')
            return False


    def reset(self):
        try:
            self.page.click("xpath=//*[@id='breadcrumbs']/a[2]", timeout=5000)

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
            icon = self.page.wait_for_selector("xpath=//*[@id='logo']/a/img", timeout=1000)
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
            print(f"🚨🚨 An unexpected error occurred: {str(e)[:100]}\nErro name:{type(e).__name__}")
            return False


    def requirements(self):
        txt_ = (f'Para atendimento integral da diligência inserida na aba "Pareceres" em '
                f'{datetime.now().strftime("%d/%m/%Y")}.')
        # Select Rquisites
        self.page.click("xpath=//div[@id='div_2144784112']//span//span[contains(text(),'Requisitos')]"
                        , timeout=5000)
        #
        self.page.click("xpath=//span[contains(text(),'Requisitos para Celebração')]",
                        timeout=5000)
        #
        self.page.click("xpath=//input[@id='formRequisitosDocumentos:_idJsp159']", timeout=5000)
        #
        txt_field = "xpath=//textarea[@id='formSolicitacaComplementacao:observacao']"
        self.page.fill(selector=txt_field, value=f"{txt_}", timeout=1000)

        self.page.click("xpath=//input[@id='formSolicitacaComplementacao:_idJsp23']")

        complete_ = self.page.wait_for_selector("xpath=//div[@class='messages']//div", timeout=5000)
        if complete_:
            print("✅🤝 Solicitação de complementação enviada com sucesso")
            self.logger.info('Solicitação de complementação enviada com sucesso')

    def error_message(self):
        try:
            self.page.wait_for_selector('//*[@id="popUpLayer2"]', timeout=2000)
            close_button = self.page.locator('#popUpLayer2 img[src*="close.gif"]')
            close_button.wait_for(state="visible", timeout=5000)
            close_button.click()
        except Exception as e:
            print(f"🚨🚨 An unexpected error occurred: {str(e)[:100]}\nErro name:{type(e).__name__}")
            self.logger.info('Erro ao solicitar complementação')
            return False



# convenio
def main():
    text_conv = r'''Prezados,

Favor desconsiderar a diligência anterior (08/10/2025), tendo em vista o equívoco relacionado ao link 2. 

Dessa forma, com vistas à celebração da parceria e conforme exigido nas normativas legais vigente, solicitamos o encaminhamento da seguinte documentação:

1. Cópia da LOA (Lei Orçamentária Anual);
2. Cópia do QDD (Quadro de Detalhamento de Despesa);
3. Documento que apresente o número da matrícula funcional do Representante Legal; e
4. Termo de Posse/Nomeação ou Diploma do Representante Legal.
5. Declarações Consolidadas – Sem prazo de validade (Link 1); e
6. Declarações Consolidadas – Validade no mês da assinatura (Link 2).

Link 1 – Declarações Consolidadas – Sem prazo de validade:
https://sneaelis.itech.ifsertaope.edu.br/forms/declaracoes/

Link 2 – Declarações Consolidadas – Validade no mês da assinatura:
https://homologacao.itech.ifsertaope.edu.br/forms/declaracoes/Validade-declaracao-mes 

Após a inserção integral dessas documentações na aba “Requisitos para Celebração” do Transferegov, o Proponente deverá acionar a opção “Enviar para Análise”, disponível ao final da página.

Cumprida integralmente esta diligência, prosseguiremos com as etapas necessárias para celebração do convênio sob a condição suspensiva, cabendo ao Proponente apresentar o Termo de Referência e documentos correlatos (Projeto Técnico, Cotações e Planilha de Custos) no prazo de 09 (nove) meses a contar da assinatura do instrumento, conforme determina a legislação vigente.

Cabe destacar que, no ato da celebração da parceria, o Proponente deverá estar adimplente junto aos sistemas CAUC e Regularidade Transferegov, no que couber. Caso seja constatado qualquer registro de inadimplência, a celebração da parceria ficará inviabilizada.

Além disso, o Proponente deverá possuir cadastro de usuário externo no Sistema Eletrônico de Informações (SEI) junto ao Ministério do Esporte, para possibilitar a assinatura do instrumento. Caso o Representante Legal ainda não possua cadastro, este deverá ser realizado pelo seguinte link:

https://sei.cidadania.gov.br/sei/controlador_externo.php?acao=usuario_externo_logar&id_orgao_acesso_externo=0 

Após o cadastro, o dirigente da entidade deverá enviar a relação da documentação necessária, por meio do Protocolo Digital do Ministério do Esporte, para fins de ativação do acesso, conforme mensagem automática enviada ao e-mail vinculado ao cadastro.

Observação: Esta é uma mensagem automática. Favor desconsiderar caso a documentação supramencionada tenha sido apresentada integralmente antes da emissão deste Parecer.

Permanecemos à disposição pelo endereço eletrônico: cgfp.sneaelis@esporte.gov.br. 

Atenciosamente,

Coordenação-Geral de Formalização de Parcerias

'''

    text_tf = r'''Prezados,

Ao cumprimentá-los cordialmente e com vistas à celebração da parceria, conforme exigido nas normativas legais vigente, solicitamos o encaminhamento da seguinte documentação:

1. Ata de Eleição do Corpo de Dirigentes atual da Entidade, registrada em Cartório;
2. Termo de Posse/Nomeação do Representante Legal da Entidade;
3. Comprovante de Inscrição no CNPJ;
4. Comprovante de Endereço da Entidade atualizado;
5. Estatuto Social da Entidade;
6. Alterações Estatuárias, se for o caso, e Ata que o aprovou, registrados em Cartório;
7. Certidão Negativa de Débitos Trabalhista – CNDT (Link 1); 
8. Declarações Consolidadas (Link 2); e
9. Declaração do art. 26 e 27 (Link 3).

Link 1 – CNDT:
https://cndt-certidao.tst.jus.br/inicio.faces 

Link 2 - Declarações Consolidadas:
https://sneaelis.itech.ifsertaope.edu.br/forms/declaracoes/formulario-documentacoes 

Link 3 – Declaração do art. 26 e 27 do Decreto nº 8.726/2016:
https://sneaelis.itech.ifsertaope.edu.br/forms/declaracoes/formulario-dirigente 

Ressalta-se que a Declaração do art. 26 e 27 (Link 3), deve constar a relação nominal atualizada dos dirigentes da entidade, conforme estabelecido no estatuto, com os dados específicos de cada um deles.

Após a inserção integral dessas documentações na aba “Requisitos para Celebração” do Transferegov, o Proponente deverá acionar a opção “Enviar para Análise”, disponível ao final da página.

Cumprida integralmente esta diligência, prosseguiremos com as etapas necessárias para formalização, conforme determina a legislação vigente.

Cabe destacar que no ato da celebração da parceria, a entidade deverá estar adimplente junto aos sistemas: CAUC, CEIS/CEPIM e Regularidade Transferegov. Caso seja constatado qualquer registro de inadimplência, a celebração da parceria ficará inviabilizada.

Caberá a entidade ainda, atentar-se às vedações dispostas no art. 39, da Lei nº 13.019/2014, especialmente no dever de prestar contas, tendo em vista que se constatado pendência, a celebração da parceria ficará condicionada à devida regularização.

Além disso, o Proponente deverá possuir cadastro de usuário externo no Sistema Eletrônico de Informações (SEI) junto ao Ministério do Esporte, para possibilitar a assinatura do instrumento. Caso o Representante Legal ainda não possua cadastro, este deverá ser realizado pelo seguinte link:

https://sei.cidadania.gov.br/sei/controlador_externo.php?acao=usuario_externo_logar&id_orgao_acesso_externo=0 

Após o cadastro, o dirigente da entidade deverá enviar a relação da documentação necessária, por meio do Protocolo Digital do Ministério do Esporte, para fins de ativação do acesso, conforme mensagem automática enviada ao e-mail vinculado ao cadastro.

Observação: Esta é uma mensagem automática. Favor desconsiderar caso a documentação supramencionada tenha sido apresentada integralmente antes da emissão deste Parecer.

Permanecemos à disposição por meio do endereço eletrônico: cgfp.sneaelis@esporte.gov.br. 

Atenciosamente,

Coordenação-Geral de Formalização de Parcerias
'''

    xlsx_source_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                        r'Social\Teste001\fabi_DFP\Propostas Para Diligências Padrão.xlsm')
    # Get excel file
    excel_file = pd.ExcelFile(xlsx_source_path)
    # Get all sheets inside the file and store in a list
    sheet_names = excel_file.sheet_names

    # Initiate automation instance
    robo = PWRobo(text_tf=text_tf, text_conv=text_conv)

    robo.land_page()
    robo.init_search()

    for i, sheet in enumerate(sheet_names):
        if sheet == 'Planilha' or  sheet == 'Termo de Fomento':
            continue
        print(f"\n{'<' * 3}📄 Loading sheet [{i}/{len(sheet_names)}]: '{sheet}'{'>' * 3}"
              .center(80, '-'))
        try:
            # Create the DataFrame with source data
            df = pd.read_excel(xlsx_source_path, dtype=str, sheet_name=sheet)
            print(f"\n✅ Sheet '{sheet}' loaded successfully with {len(df)} rows.\n")

            for idx, row in df.iterrows():
                if idx < 314:
                    continue
                try:
                    porp_num = row['Nº Proposta']
                    if pd.isna(porp_num):
                        continue
                    print(f"\n{'⚡' * 3}🚀 EXECUTING PROPOSAL: {porp_num}, index: {idx} 🚀{'⚡' * 3}".center(70,
                           '='), '\n')

                    robo.loop_search(porp_num)
                    feedback_ = robo.insert_feedback(type_txt=i)
                    if feedback_:
                        robo.requirements()
                    robo.reset()
                    try:
                        if not robo.error_message():
                            robo.land_page()
                            robo.init_search()
                    except Exception as e:
                        print(
                           f"🚨🚨 An unexpected error occurred: {str(e)[:100]}\nErro name:{type(e).__name__}")

                except Exception as e:
                    print(f"🚨🚨 An unexpected error occurred: {str(e)[:100]}\nErro name:{type(e).__name__}")

        except Exception as e:
            print(f"🚨🚨 Failed to load or process sheet '{sheet}': {str(e)[:100]} (Error: {type(e).__name__})")

# termo de fomento
def main2():
    text_conv = r'''Prezados,

Favor desconsiderar a diligência anterior (08/10/2025), tendo em vista o equívoco relacionado ao link 2. 

Dessa forma, com vistas à celebração da parceria e conforme exigido nas normativas legais vigente, solicitamos o encaminhamento da seguinte documentação:

1. Cópia da LOA (Lei Orçamentária Anual);
2. Cópia do QDD (Quadro de Detalhamento de Despesa);
3. Documento que apresente o número da matrícula funcional do Representante Legal; e
4. Termo de Posse/Nomeação ou Diploma do Representante Legal.
5. Declarações Consolidadas – Sem prazo de validade (Link 1); e
6. Declarações Consolidadas – Validade no mês da assinatura (Link 2).

Link 1 – Declarações Consolidadas – Sem prazo de validade:
https://sneaelis.itech.ifsertaope.edu.br/forms/declaracoes/

Link 2 – Declarações Consolidadas – Validade no mês da assinatura:
https://homologacao.itech.ifsertaope.edu.br/forms/declaracoes/Validade-declaracao-mes 

Após a inserção integral dessas documentações na aba “Requisitos para Celebração” do Transferegov, o Proponente deverá acionar a opção “Enviar para Análise”, disponível ao final da página.

Cumprida integralmente esta diligência, prosseguiremos com as etapas necessárias para celebração do convênio sob a condição suspensiva, cabendo ao Proponente apresentar o Termo de Referência e documentos correlatos (Projeto Técnico, Cotações e Planilha de Custos) no prazo de 09 (nove) meses a contar da assinatura do instrumento, conforme determina a legislação vigente.

Cabe destacar que, no ato da celebração da parceria, o Proponente deverá estar adimplente junto aos sistemas CAUC e Regularidade Transferegov, no que couber. Caso seja constatado qualquer registro de inadimplência, a celebração da parceria ficará inviabilizada.

Além disso, o Proponente deverá possuir cadastro de usuário externo no Sistema Eletrônico de Informações (SEI) junto ao Ministério do Esporte, para possibilitar a assinatura do instrumento. Caso o Representante Legal ainda não possua cadastro, este deverá ser realizado pelo seguinte link:

https://sei.cidadania.gov.br/sei/controlador_externo.php?acao=usuario_externo_logar&id_orgao_acesso_externo=0 

Após o cadastro, o dirigente da entidade deverá enviar a relação da documentação necessária, por meio do Protocolo Digital do Ministério do Esporte, para fins de ativação do acesso, conforme mensagem automática enviada ao e-mail vinculado ao cadastro.

Observação: Esta é uma mensagem automática. Favor desconsiderar caso a documentação supramencionada tenha sido apresentada integralmente antes da emissão deste Parecer.

Permanecemos à disposição pelo endereço eletrônico: cgfp.sneaelis@esporte.gov.br. 

Atenciosamente,

Coordenação-Geral de Formalização de Parcerias

'''

    text_tf = r'''Prezados,

Ao cumprimentá-los cordialmente e com vistas à celebração da parceria, conforme exigido nas normativas legais vigente, solicitamos o encaminhamento da seguinte documentação:

1. Ata de Eleição do Corpo de Dirigentes atual da Entidade, registrada em Cartório;
2. Termo de Posse/Nomeação do Representante Legal da Entidade;
3. Comprovante de Inscrição no CNPJ;
4. Comprovante de Endereço da Entidade atualizado;
5. Estatuto Social da Entidade;
6. Alterações Estatuárias, se for o caso, e Ata que o aprovou, registrados em Cartório;
7. Certidão Negativa de Débitos Trabalhista – CNDT (Link 1); 
8. Declarações Consolidadas (Link 2); e
9. Declaração do art. 26 e 27 (Link 3).

Link 1 – CNDT:
https://cndt-certidao.tst.jus.br/inicio.faces 

Link 2 - Declarações Consolidadas:
https://sneaelis.itech.ifsertaope.edu.br/forms/declaracoes/formulario-documentacoes 

Link 3 – Declaração do art. 26 e 27 do Decreto nº 8.726/2016:
https://sneaelis.itech.ifsertaope.edu.br/forms/declaracoes/formulario-dirigente 

Ressalta-se que a Declaração do art. 26 e 27 (Link 3), deve constar a relação nominal atualizada dos dirigentes da entidade, conforme estabelecido no estatuto, com os dados específicos de cada um deles.

Após a inserção integral dessas documentações na aba “Requisitos para Celebração” do Transferegov, o Proponente deverá acionar a opção “Enviar para Análise”, disponível ao final da página.

Cumprida integralmente esta diligência, prosseguiremos com as etapas necessárias para formalização, conforme determina a legislação vigente.

Cabe destacar que no ato da celebração da parceria, a entidade deverá estar adimplente junto aos sistemas: CAUC, CEIS/CEPIM e Regularidade Transferegov. Caso seja constatado qualquer registro de inadimplência, a celebração da parceria ficará inviabilizada.

Caberá a entidade ainda, atentar-se às vedações dispostas no art. 39, da Lei nº 13.019/2014, especialmente no dever de prestar contas, tendo em vista que se constatado pendência, a celebração da parceria ficará condicionada à devida regularização.

Além disso, o Proponente deverá possuir cadastro de usuário externo no Sistema Eletrônico de Informações (SEI) junto ao Ministério do Esporte, para possibilitar a assinatura do instrumento. Caso o Representante Legal ainda não possua cadastro, este deverá ser realizado pelo seguinte link:

https://sei.cidadania.gov.br/sei/controlador_externo.php?acao=usuario_externo_logar&id_orgao_acesso_externo=0 

Após o cadastro, o dirigente da entidade deverá enviar a relação da documentação necessária, por meio do Protocolo Digital do Ministério do Esporte, para fins de ativação do acesso, conforme mensagem automática enviada ao e-mail vinculado ao cadastro.

Observação: Esta é uma mensagem automática. Favor desconsiderar caso a documentação supramencionada tenha sido apresentada integralmente antes da emissão deste Parecer.

Permanecemos à disposição por meio do endereço eletrônico: cgfp.sneaelis@esporte.gov.br. 

Atenciosamente,

Coordenação-Geral de Formalização de Parcerias
'''

    xlsx_source_path = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
                        r'Social\Teste001\fabi_DFP\Propostas Para Diligências Padrão.xlsm')
    # Get excel file
    excel_file = pd.ExcelFile(xlsx_source_path)
    # Get all sheets inside the file and store in a list
    sheet_names = excel_file.sheet_names

    # Initiate automation instance
    robo = PWRobo(text_tf=text_tf, text_conv=text_conv)

    robo.land_page()
    robo.init_search()

    for i, sheet in enumerate(sheet_names):
        if sheet == 'Planilha' or  sheet == 'Convênio':
            continue
        print(f"\n{'<' * 3}📄 Loading sheet [{i}/{len(sheet_names)}]: '{sheet}'{'>' * 3}"
              .center(80, '-'))
        try:
            # Create the DataFrame with source data
            df = pd.read_excel(xlsx_source_path, dtype=str, sheet_name=sheet)
            print(f"\n✅ Sheet '{sheet}' loaded successfully with {len(df)} rows.\n")

            for idx, row in df.iterrows():
                try:
                    porp_num = row['Nº Proposta']
                    if pd.isna(porp_num):
                        continue
                    print(f"\n{'⚡' * 3}🚀 EXECUTING PROPOSAL: {porp_num} 🚀{'⚡' * 3}".center(70, '=')
                          , '\n')
                    robo.loop_search(porp_num)
                    robo.requirements()
                    robo.reset()
                    try:
                        if not robo.error_message():
                            robo.land_page()
                            robo.init_search()
                    except Exception as e:
                        print(
                           f"🚨🚨 An unexpected error occurred: {str(e)[:100]}\nErro name:{type(e).__name__}")

                except Exception as e:
                    print(f"🚨🚨 An unexpected error occurred: {str(e)[:100]}\nErro name:{type(e).__name__}")

        except Exception as e:
            print(f"🚨🚨 Failed to load or process sheet '{sheet}': {str(e)[:100]} (Error: {type(e).__name__})")


if __name__ == "__main__":
    main()
    time.sleep(5)
    main2()



#25448_2025