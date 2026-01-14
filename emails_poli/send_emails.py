import os
import sys

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

            print(f'{len(correct_email_xlsx)} and {len(final_email_set)}')

            return final_email_set

        except Exception as e:
            print(f"❌ Failed to read Excel file: \n{e}")
            return []


    def mark_as_done(index, file_path):
        """Safely mark row as done and save to Excel"""
        try:
            status_df = pd.read_excel(excel_path, dtype=str, sheet_name='Status')

            index = str(index)

            status_df.loc[status_df['Index'] == index, 'Status'] = 'feito'

            # Verify the modification
            modified_value = status_df.loc[status_df['Index'] == index, 'Status'].values
            if len(modified_value) > 0 and modified_value[0] == 'feito':
                print("✅ DataFrame modification successful")
            else:
                print("❌ DataFrame modification failed!")

            with ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # Save to Excel
                status_df.to_excel(writer, index=False, sheet_name='Status')

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
            e_mail.Send()
            return print(f"📧📤 Email sent to: {email}")
        except Exception as e:
            print(f"❌ Failed to send email: \n{e}")
            return False


    def create_email_content():
        """
        Customize this function to generate your email content
        based on the Excel data for each recipient
        """
        set_subject = (f"Canal Oficial da SNEAELIS no WhatsApp - Link de Acesso")

        set_html_body = f"""
<p>Prezados(as),</p>

<p>A Secretaria Nacional de Esporte Amador, Educação, Lazer e Inclusão Social (SNEAELIS) está disponibilizando um canal oficial no WhatsApp, criado para o envio de avisos e comunicados importantes, de forma mais ágil e organizada.</p>

<p>Para acessar o grupo, utilize o link abaixo:</p>

<p><a href="https://chat.whatsapp.com/IVun6wpHIcD8TlKlQeamJu?mode=hqrt1">
👉 Clique aqui para acessar o canal oficial no WhatsApp</a></p>

<p>O canal é destinado exclusivamente à divulgação de informações oficiais. Em caso de dúvidas ou demandas, os atendimentos seguem pelos meios institucionais habituais.</p>

<p>Agradecemos a atenção e contamos com a participação de todos.</p>

<p>Atenciosamente,<br>
Secretaria Nacional de Esporte Amador, Educação, Lazer e Inclusão Social – SNEAELIS<br>
Ministério do Esporte</p>
"""
        return set_subject, set_html_body

    email_data = read_excel_data()

    if not email_data:
        print("No data found in Excel file")
        return

    count = 0
    for index, email in enumerate(email_data):
        if not email:
            continue
        #entidade = entidade_data[index]
        attachment = attachment_path
        subject, html_body = create_email_content()

        # Send email
        send_email_outlook()
        #mark_as_done(index=index, file_path=excel_path)
        count += 1

    print(count)


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
    xlsx = r"C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência Social\Teste001\propostas_sneaelis.xlsx"
    #attach = (r'')
    sender = 'assessoria.sneaelis@esporte.gov.br'

    send_emails_from_excel(
        excel_path=xlsx,
        attachment_path='',
        sender=sender
    )


if __name__ == "__main__":
    send_emails_from_excel_main()
    #Try to commit