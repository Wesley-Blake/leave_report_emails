from pathlib import Path

import pandas as pd
import validators
from win32com_email import win_email


def month_selection() -> str:
    months_dict = {
        "1": "January",
        "2": "February",
        "3": "March",
        "4": "April",
        "5": "May",
        "6": "June",
        "7": "July",
        "8": "August",
        "9": "September",
        "10": "October",
        "11": "November",
        "12": "December",
    }
    while True:
        for i, value in months_dict.items():
            print(i, value)
        month = input("Month of interest (int): ")
        if input(f"Is {months_dict[month]} correct? [Y/n] ") == "n":
            continue
        if month.isdigit() and month in months_dict:
            return months_dict[month]
        else:
            print("Failed to detect Month.")


def loading_bar(length, index=1, pre_fix=""):
    BAR_LENGTH = 30
    print()
    if len(pre_fix) > 0:
        print(pre_fix)
    while index <= length:
        block = int(BAR_LENGTH * index / length)
        bar = "=" * block + "-" * (BAR_LENGTH - block)
        yield (f"\r|{bar}| {index} / {length} emails sent.")
        index += 1


def leave_reports(file: Path | str) -> dict[str, str]:
    if not isinstance(file, Path):
        file = Path(file)
    if not file.is_file():
        return {}
    df = pd.read_csv(file)
    if df.empty:
        print("Empty csv.")
        return {}

    white_list = [
        "EmplID",
        "FirstName",
        "LastName",
        "leave_report_status",
        "ApproverEmail",
        "EmplEmail",
    ]
    df = df[white_list]
    df["EmplID"] = df["EmplID"].astype(str)
    df["Combined"] = df["EmplID"] + " " + df["FirstName"] + " " + df["LastName"]
    filtered_df = df[~(df["leave_report_status"] == "Completed")]

    result = {}
    manager = df["ApproverEmail"].unique().tolist()
    for email_m in manager:
        if not validators.email(email_m):
            print(f"Skipping row(s) with invalid ApproverEmail: {email_m!r}")
            continue
        employee_list = (
            filtered_df[filtered_df["ApproverEmail"] == email_m]["Combined"]
            .unique()
            .tolist()
        )
        if len(employee_list) > 0:
            result.update({email_m: employee_list})

    return result


if __name__ == "__main__":
    DOWNLOADS = Path.home() / "Downloads"

    # PAY_MONTH = input('Enter Month: ')
    PAY_MONTH = month_selection()

    files = []
    if not DOWNLOADS.is_dir():
        raise SystemExit("Downloads path doesn't exist.")
    for file in DOWNLOADS.iterdir():
        if file.name.startswith("Leave_Report_Status"):
            files.append(file.absolute())
    target = max(files)
    emails = leave_reports(target)
    length = len(emails)
    my_bar = loading_bar(length, pre_fix="Manager Emails:")
    for manager, employees in emails.items():
        body = f"""
Hi,

The below Leave Reports aren't complete yet.
The month of {PAY_MONTH} is due on the 15th of the following month.

*** If you see a person that shouldn't be there or multiple Leave
Reports for 1 person, let me know. ***
"""
        body += "\n".join(employees)
        print(next(my_bar), end="", flush=True)
        win_email(manager, [], PAY_MONTH, body)
    print()
