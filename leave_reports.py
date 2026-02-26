from pathlib import Path
import pandas as pd
from win32com_email import email

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
        "12": "December"
    }
    while True:
        for i, value in months_dict.items():
            print(i,value)
        month = input("Month of interest (int): ")
        if month.isdigit() and month in months_dict.keys():
            return months_dict[month]
        else:
            print("Failed to detect Month.")


def leave_reports(file: str) -> dict[str,str]:
    file = Path(file)
    result = {}

    if file.is_file():
        df = pd.read_csv(file)
        if df.empty:
            print("Empty csv.")
            return {}

        white_list = ['EmplID','FirstName','LastName','leave_report_status', 'ApproverEmail', 'EmplEmail']
        df = df[white_list]
        df['EmplID'] = df['EmplID'].astype(str)
        #df['Combined'] = df[['EmplID','FirstName','LastName']].astype(str).agg(' '.join, axis=1)
        df['Combined'] = df['EmplID'] + ' ' + df['FirstName'] + ' ' + df['LastName']
        filtered_df = df[~(df['leave_report_status'] == 'Completed')]
    else:
        return {}

    manager = df['ApproverEmail'].unique().tolist()
    for email in manager:
        employee_list = filtered_df[
            filtered_df['ApproverEmail'] == email
        ]['Combined'].unique().tolist()
        if len(employee_list) > 0:
            result.update({email: employee_list})

    return result


if __name__ == '__main__':
    DOWNLOADS = Path.home() / 'Downloads'

    #PAY_MONTH = input('Enter Month: ')
    PAY_MONTH = month_selection()

    files = []
    if DOWNLOADS.is_dir():
        for file in DOWNLOADS.iterdir():
            if file.name.startswith('Leave_Report_Status'):
                files.append(file.absolute())
        target = max(files)
        emails = leave_reports(target)
        for manager, employees in emails.items():
            body = \
f"""
Hi,

The below Leave Reports aren't complete yet.
The month of {PAY_MONTH} is due on the 15th of the following month.

*** If you see a person that shouldn't be there or multiple Leave Reports for 1 person, let me know. ***
"""
            body += "\n".join(employees)
            email(manager, [], PAY_MONTH, body)

