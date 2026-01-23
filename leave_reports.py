from pathlib import Path
import pandas as pd
from win32com_email import email

def leave_reports(file: str) -> dict[str,str]:
    file = Path(file)
    result = {}

    if file.is_file():
        df = pd.read_csv(file)

        white_list = ['EmplID','FirstName','LastName','leave_report_status', 'ApproverEmail', 'EmplEmail']
        df = df[white_list]
        df['Combined'] = df[['A', 'B', 'C']].agg(''.join, axis=1)
        filtered_df = df[~(df['leave_report_status'] == 'Completed')]
    else:
        return {}

    manager = df['ApproverEmail'].unique().tolist()
    for email in manager:
        employee_list = filtered_df[
            filtered_df['ApproverEmail'] == email
        ]['EmplEmail'].unique().tolist()
        if len(employee_list) > 0:
            result.update({email: employee_list})

    return result


if __name__ == '__main__':
    DOWNLOADS = Path.home() / 'Downloads'

    PAY_MONTH = input('Enter Month: ')

    files = []
    if DOWNLOADS.is_dir():
        for file in DOWNLOADS.iterdir():
            if file.name.startswith('Leave_Report_Status'):
                files.append(file.absolute())
        target = max(files)
        emails = leave_reports(target)
        for manager, employees in emails.items():
            email(manager, [], PAY_MONTH,\
f"""
Hi,

The bellow are incomplete as of {PAY_MONTH}.
{employees}
"""
)
