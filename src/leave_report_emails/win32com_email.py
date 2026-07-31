import configparser
from pathlib import Path

import win32com.client as win32

cfg = configparser.ConfigParser()
cfg.read(Path(__name__).parent / ".env")
LINK = Path(cfg["leave_report_emails"]["how_to_submit"])


def win_email(cc: str, bcc: list[str], pay_period: str, body: str) -> None:
    try:
        outlook = win32.Dispatch("outlook.application")
    except Exception as e:
        raise SystemError("Failed to create Outlook application") from e

    mail = outlook.CreateItem(0)
    mail.CC = cc
    mail.BCC = "; ".join(bcc)
    mail.Subject = f"Leave Report Month {pay_period}"
    if LINK.is_file():
        mail.Attachments.Add(str(LINK))
    mail.Body = body
    # mail.Display()
    mail.Send()
