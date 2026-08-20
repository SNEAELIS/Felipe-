import os
import sys
import time
from datetime import datetime

import win32com.client as win32
from jinja2 import Template


def send_daily_briefing(
    politician_name: str,
    grouped_news: dict,
    recipient_email: str,
):
    """Renders the HTML newsletter and dispatches it via Outlook.

    Opens an Outlook compose window (Display) so the sender can review and
    click Send manually — mirrors emails_poli/send_emails.py.

    PRODUCTION NOTE: replace `e_mail.Display()` with `e_mail.Send()` and
    remove the `sys.exit()` / `time.sleep()` calls to auto-send.
    """

    # Load and render Jinja2 template
    template_path = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "newsletter_template.html"
    )
    with open(template_path, "r", encoding="utf-8") as f:
        template = Template(f.read())

    html_content = template.render(
        politician_name=politician_name,
        date_today=datetime.now().strftime("%d/%m/%Y"),
        grouped_news=grouped_news,
    )

    subject = f"[Briefing 24h] {politician_name} - {datetime.now().strftime('%d/%m')}"

    # Send email through Outlook
    outlook = win32.Dispatch("outlook.application")
    e_mail = outlook.CreateItem(0)  # 0 = olMailItem
    e_mail.SentOnBehalfOfName = "felipe.rsouza@esporte.gov.br"
    e_mail.To = recipient_email
    e_mail.Subject = subject
    e_mail.HTMLBody = html_content

    # Move from Drafts to Outbox and send
    try:
        # DEBUG: open the draft for manual review (mirrors send_emails.py)
        # PRODUCTION: change Display() -> Send() and remove sys.exit() calls
        e_mail.Display()
        sys.exit()
        time.sleep(1.5)
        return print(f"Email sent to: {recipient_email}")

    except:
        sys.exit()

