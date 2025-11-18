import os
import sys

import win32com.client as win32
from fuzzywuzzy import fuzz, process
import pandas as pd


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
        try:
            df = pd.read_excel(excel_path, dtype=str)
            # Cria um lista para cada coluna do arquivo xlsx
            email_xlsx = list()
            entidade_xlsx = list()

            # Itera a planilha e armazena os dados em listas
            for indice, linha in df.iterrows():  # Assume que a primeira linha e um cabeçalho
                email_xlsx.append(linha["Email parlamentar"])  # Busca destinatário do email
                entidade_xlsx.append(linha["Parlamentar"])  # Busca o destinatário da mensagem

            # Fix domain problems
            correct_email_xlsx = [correct_email_domain(e) for e in email_xlsx]

            return entidade_xlsx, correct_email_xlsx
        except Exception as e:
            print(f"❌ Failed to read Excel file: \n{e}")
            return []


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
            e_mail.Display()
            sys.exit()
            return print(f"📧📤 Email sent to: {email}")
        except Exception as e:
            print(f"❌ Failed to send email: \n{e}")
            return False


    def create_email_content():
        """
        Customize this function to generate your email content
        based on the Excel data for each recipient
        """
        set_subject = (f"Link - Webinar de Orientações para apresentação de propostas à SNEAELIS - "
                       f"Emendas Parlamentares 2025")

        set_html_body = f"""
        <p>Prezado(a)  {entidade},</p>

        <p>Você está recebendo esse e-mail porque se inscreveu pra o Webinar de orientações para apresentação
         de propostas à SNEAELIS - Emendas Parlamentares 2025.</p>

        <p>O evento acontece hoje, às 13h30 e você pode acompanhar através do link abaixo.</p>
        
        <p><a href="https://youtube.com/live/o3354lllE4Y?feature=share">
        Clique aqui para acessar o evento</a></p>        
        
        <p>Nos vemos em breve!</p>      
        """
        return set_subject, set_html_body

    entidade_data, email_data = read_excel_data()

    if not email_data:
        print("No data found in Excel file")
        return

    for index, email in enumerate(email_data):
        entidade = entidade_data[index]
        attachment = attachment_path
        subject, html_body = create_email_content()

        # Send email
        send_email_outlook()


def send_emails_from_excel_main():
    xlsx = (r'C:\Users\felipe.rsouza\OneDrive - Ministério do Desenvolvimento e Assistência '
            r'Social\Teste001\emails_poli\Formulário de Inscrição-Webinário 2025.xlsx')
    #attach = (r'')
    sender = 'assessoria.sneaelis@esporte.gov.br'

    send_emails_from_excel(
        excel_path=xlsx,
        attachment_path='',
        sender=sender
    )


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
    if correct_domains is None:
        correct_domains = [
            'gmail.com', 'yahoo.com', 'outlook.com', 'hotmail.com',
            'aol.com', 'icloud.com', 'protonmail.com', 'mail.com',
            'gmail', 'yahoo', 'outlook', 'hotmail'  # Also allow bare domains
        ]

    # Extract the domain part
    parts = email.split('@')
    if len(parts) != 2:
        return email  # Not a valid email format

    local_part, domain = parts

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


if __name__ == "__main__":
    send_emails_from_excel_main()
    #Try to commit
