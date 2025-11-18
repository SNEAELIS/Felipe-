# Email Sender & Domain Fixer Tool

## Overview

This tool automates the process of sending emails with attachments to a list of recipients whose information is stored in an `.xlsx` file. It is especially useful for cases where some email addresses may have incorrect or outdated domains. The script cross-references a domain correction list (defined in the code) and automatically fixes invalid domains before sending out the emails.

## Features

- **Reads Recipients from Excel**: Loads names and email addresses from an Excel spreadsheet.
- **Domain Correction**: Checks each email's domain against a list of known valid domains and corrects them if mismatches are found.
- **Automated Email Sending**: Sends a predefined message and optional attachment to each corrected recipient address.
- **Logging**: Outputs the status of each email sent (success or failure) to the console.

## How It Works

1. **Input Excel File**: The tool expects an `.xlsx` file containing at least two columns: `Name` and `Email`.
2. **Domain Correction List**: Inside the script, a dictionary maps common incorrect domains to their correct versions.
3. **Processing**: Each email is checked and corrected if needed.
4. **Personalized Sending**: An email is composed and sent to each recipient, optionally including an attachment.

## Usage

1. Place your `.xlsx` file in the script directory.
2. Adjust the script's configuration section with your message, attachment path, and domain correction dictionary.
3. Run the script. It will:
    - Load the data
    - Fix any incorrect domains
    - Send the emails with the message and attachment

## Example

Suppose your `recipients.xlsx` looks like this:

| Name         | Email                    |
|--------------|--------------------------|
| Ana Pereira  | ana.pereira@gmai.com     |
| João Silva   | joao.silva@outlok.com    |
| Maria Souza  | maria.souza@esporte.gov.br |

The domain correction dictionary in the script:
```python
domain_corrections = {
    "gmai.com": "gmail.com",
    "outlok.com": "outlook.com",
    # Add more corrections as needed
}
```

After processing, the script will send emails to:
- ana.pereira@gmail.com
- joao.silva@outlook.com
- maria.souza@esporte.gov.br

## Requirements

- Python 3.x
- pandas
- openpyxl
- win32com (for Outlook integration)

Install dependencies:
```sh
pip install pandas openpyxl pywin32
```

## Configuration

- **Message**: Edit the message text inside the script.
- **Attachment**: Set the path to your attachment file.
- **Domain Corrections**: Update the dictionary as needed.

## Disclaimer

- You must have Outlook installed and configured on your machine.
- Make sure you have permissions to send emails from your Outlook profile.

## Sample Code

```python
import pandas as pd
import win32com.client as win32

# Domain corrections dictionary
domain_corrections = {
    "gmai.com": "gmail.com",
    "outlok.com": "outlook.com",
    # Add more as needed
}

def correct_email(email):
    if '@' in email:
        user, domain = email.split('@')
        domain = domain_corrections.get(domain, domain)
        return f"{user}@{domain}"
    return email

def send_emails_from_excel(xlsx_path, subject, body, attachment=None):
    df = pd.read_excel(xlsx_path)
    outlook = win32.Dispatch('outlook.application')
    for idx, row in df.iterrows():
        name = row['Name']
        email = correct_email(row['Email'])
        mail = outlook.CreateItem(0)
        mail.To = email
        mail.Subject = subject
        mail.Body = body
        if attachment:
            mail.Attachments.Add(attachment)
        try:
            mail.Send()
            print(f"Sent to {name} <{email}>")
        except Exception as e:
            print(f"Failed to send to {name} <{email}>: {e}")

# Example usage
send_emails_from_excel(
    xlsx_path="recipients.xlsx",
    subject="Important Notice",
    body="Dear recipient, please see the attached document.",
    attachment="document.pdf"
)
```

---

Feel free to adapt the script to your specific needs!