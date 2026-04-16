import random
import os
import sys
import time

import win32com.client as win32
from fuzzywuzzy import fuzz, process
import pandas as pd
from pandas import ExcelWriter




# Send email to a list that is on excel, email is sent changing the sender and also has text and attachemnt
def send_emails_from_excel(excel_path, attachment_path=None, sender=None):
    """
    Main function to process Excel and send emails

    Args:
        excel_path (str): Path to Excel file with recipient data
        attachment_path (string): Attachment file paths (optional)
        sender (str): Email address to send from (optional)
    """

    sent_mails = []

    def read_excel_data():
        """
        Read email data from Excel file
        Returns: List
        """
        unique_emails = set()
        correct_email_xlsx = list()
        final_email_set = set()

        try:
            excel_file = pd.ExcelFile(excel_path)
            sheets_to_process = excel_file.sheet_names
            print(sheets_to_process)

            for sheet in sheets_to_process:
                print(f"Processing sheet: {sheet}")

                df = pd.read_excel(excel_path, dtype=str, sheet_name=sheet)
                if 'Proponente' in df.columns:
                    column_values = df['Proponente'].dropna().astype(str).str.strip()
                    unique_emails.update(column_values)

                print(f'total data_frame size: {len(df)}\nunique values: {len(unique_emails)}')

            # Fix domain problems
            for e in unique_emails:
                if not '@' in e:
                    continue
                if ' ' in e :
                    e = e.split(' ')[-1]
                correct_email_xlsx.append(correct_email_domain(e))

            for i in correct_email_xlsx:
                final_email_set.add(i)

            print(f'Final number of unique emails found {len(final_email_set)}')

            return final_email_set

        except Exception as e:
            print(f"❌ Failed to read Excel file: \n{e}")
            return []


    def mark_as_done(proponente, file_path):
        """Safely mark row as done and save to Excel"""
        try:
            status_df = pd.read_excel(file_path, dtype=str)

            if 'Status' not in status_df.columns:
                status_df['Status'] = ''

            proponente = str(proponente)

            status_df.loc[status_df['Proponente'] == proponente, 'Status'] = 'feito'

            # Verify the modification
            modified_value = status_df.loc[status_df['Proponente'] == proponente, 'Status'].values
            if len(modified_value) > 0 and modified_value[0] == 'feito':
                print("✅ DataFrame modification successful")
            else:
                print("❌ DataFrame modification failed!")

            with ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # Save to Excel
                status_df.to_excel(writer, index=False)

            return True

        except Exception as err:
            print(f"❌ Error in mark_as_done: {str(err)[:100]}\n")
            return False


    def send_email_outlook():
        # Send email through Outlook with attachment
        try:
            outlook = win32.Dispatch('outlook.application')
            e_mail = outlook.CreateItem(0)
            # Set sender if specified (must be configured in Outlook)
            if sender:
                e_mail.SentOnBehalfOfName = sender

            e_mail.To = email
            e_mail.Subject = subject
            e_mail.HTMLBody = html_body

            # Add attachment
            if attachment:
                if os.path.exists(attachment):
                    e_mail.Attachments.Add(attachment)
                else:
                    print(f"⚠️ Attachment not found: {attachment}")

            # Move from Drafts to Outbox and send
            try:
                # This sometimes bypasses security
                e_mail.Display()
                sys.exit()
                time.sleep(1.5)
                return print(f"📧📤 Email sent to: {email}")

            except:
                print(sent_mails)
                sys.exit()

        except Exception as e:
            print(f"❌ Failed to send email: \n{e}")
            return False


    def create_email_content():
        """
        Customize this function to generate your email content
        based on the Excel data for each recipient
        """
        set_subject = (f"Atualização no e-SNEAELIS")

        set_html_body = f"""
<p>📢✨ Início das etapas das propostas 2026</p>

<p>Chegou o momento de dar início às etapas de preenchimento dos projetos e organização da documentação das propostas para o exercício de 2026. 📑🚀</p>

<p>As entidades que possuem Convênio ou Termo de Fomento para 2026 já podem acessar o e-SNEAELIS e iniciar o preenchimento da proposta no sistema. 💻📊</p>

<p>👉 Acesso ao sistema:</p>
<p><a href="https://sneaelis.app.br">👉 https://sneaelis.app.br</a></p>

<p>Este é o primeiro passo para avançarmos juntos na construção e formalização dos projetos. Quanto antes o preenchimento for iniciado, mais ágil será o andamento das próximas etapas. ⏳✔️</p>

<p>✨ É hora de começar!</p>

<p>Com organização e atenção aos detalhes, seguimos avançando para que as propostas evoluam com segurança e eficiência. 💙🚀</p>

<p>📢 Participe do nosso canal oficial no WhatsApp para receber comunicados, atualizações e avisos importantes:</p>

<p><a href="https://chat.whatsapp.com/IVun6wpHIcD8TlKlQeamJu?mode=hqrt1">
👉 Clique aqui para acessar o canal oficial no WhatsApp</a></p>
"""

        return set_subject, set_html_body

    email_data = read_excel_data()

    if not email_data:
        print("No data found in Excel file")
        return

    df = pd.read_excel(excel_path, dtype=str)
    if 'Status' not in df.columns:
        df['Status'] = ''
    emails_sent = df['Status'].tolist()


    count = 0
    BATCH_SIZE = 20  # Pause after this many emails
    SHORT_SLEEP = (2, 5)  # Random seconds between individual emails
    LONG_SLEEP = (60, 120)  # Random seconds to wait after a batch

    for index, email in enumerate(email_data):
        email = 'guilherme.tavares@esporte.gov.br'
        if index > 0:
            sys.exit()
        if not email or email in emails_sent:
            continue

        attachment = attachment_path
        subject, html_body = create_email_content()

        # Send email
        send_email_outlook()

        count += 1
        sent_mails.append(email)
        print(f"Sent {count}: {email}")

        # Short jitter delay to mimic human behavior
        time.sleep(random.uniform(*SHORT_SLEEP))

        # Batch delay to let the server breathe
        if count % BATCH_SIZE == 0:
            wait_time = random.randint(*LONG_SLEEP)
            print(f"Batch limit reached. Cooling down for {wait_time} seconds...")
            time.sleep(wait_time)

    print(f"Finished! Total sent: {count}")


def correct_email_domain(email, domain_threshold=80, correct_domains=None):
    """
    Corrects common email domain typos using fuzzy string matching.

    Args:
        email (str): The email address to be corrected
        domain_threshold (int): Minimum fuzzy match score to apply correction (0-100)
        correct_domains (list): List of known correct domains to match against

    Returns:
        str: The corrected email address, or original if no good match found
    """
    # Default list of common correct domains
    try:
        if correct_domains is None:
            correct_domains = [
                'gmail.com', 'yahoo.com', 'outlook.com', 'hotmail.com',
                'aol.com', 'icloud.com', 'protonmail.com', 'mail.com',
                'gmail', 'yahoo', 'outlook', 'hotmail'  # Also allow bare domains
            ]

        # Handle NaN, None, and empty values
        if pd.isna(email) or email is None or email == '':
            return None

        # Convert to string and strip whitespace
        email = str(email).strip()

        # Check for empty string after stripping
        if not email:
            return None

        # Extract the domain part
        parts = email.split('@')
        if len(parts) != 2:
            return None  # Not a valid email format

        local_part, domain = parts

        # Check if local part or domain are empty
        if not local_part or not domain:
            return None

        # Find the best match among correct domains
        best_match, score = process.extractOne(domain, correct_domains, scorer=fuzz.ratio)

        if score >= domain_threshold:
            # Reconstruct email with corrected domain
            corrected_email = f"{local_part}@{best_match}"

            # Ensure .com is present if the corrected domain is a common one
            common_domains = ['gmail', 'yahoo', 'outlook', 'hotmail']
            if best_match in common_domains:
                corrected_email = f"{local_part}@{best_match}.com"

            return corrected_email

        return email

    except Exception as e:
        # Log the error and return None for invalid emails
        print(f"Error processing email '{email}': {str(e)}")
        return None


def send_emails_from_excel_main():
    xlsx = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\emails_poli\Emails.xlsx"
    #attach = (r'')
    sender = 'felipe.rsouza@esporte.gov.br'#'assessoria.sneaelis@esporte.gov.br'

    send_emails_from_excel(
        excel_path=xlsx,
        attachment_path='',
        sender=sender
    )


if __name__ == "__main__":
    send_emails_from_excel_main()
