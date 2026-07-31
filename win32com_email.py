from pathlib import Path

import win32com.client as win32


def win_email(cc: str, bcc: list[str], pay_period: str, body: str) -> None:
    try:
        outlook = win32.Dispatch("outlook.application")
    except Exception as e:
        raise SystemError("Failed to create Outlook application") from e

    mail = outlook.CreateItem(0)
    mail.CC = cc
    mail.BCC = "; ".join(bcc)
    mail.Subject = f"Leave Report Month {pay_period}"
    with open("leave_report_emails\\secrets.txt", "r") as file:
        attachment = Path(file.readline().strip())
    if attachment.is_file():
        mail.Attachments.Add(str(attachment))
    mail.Body = body
    # mail.Display()
    mail.Send()
