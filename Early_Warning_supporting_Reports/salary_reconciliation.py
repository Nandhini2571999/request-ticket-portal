import os
import glob 
import pandas as pd 
import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
import logging 
import numpy as np 
from datetime import datetime,timedelta
from subprocess import call 
from pathlib import Path
import config 
import shutil
import re 

# Automatically get the root path (assuming script is inside your Git project)
ROOT_DIR = Path(__file__).resolve().parent
PARENT_ROOT_DIR = Path(__file__).resolve().parent.parent
REPORT_PATH = ROOT_DIR / "Reports"
OUTPUT_PATH = ROOT_DIR / "output"
log_file_path = os.path.join(ROOT_DIR, 'salary_reconcilation.log') 
# Configure logging
logging.basicConfig(
    level=logging.DEBUG,filename=log_file_path,filemode='a',  # Set the lowest level to capture
    format='%(asctime)s - %(levelname)s - %(message)s',  # Log format
)
# send mail
def send_email(to_address, subject, body_content, file_name: list = None):
    from_address = config.EMAIL_FROM
    msg = MIMEMultipart()
    from_name = "TASC GROUP"  # Customizing the From Address
    # Format the sender's address with a display name
    formatted_from = f"{from_name} <{from_address}"
    msg['From'] = formatted_from
    msg['To'] = to_address
    msg['Cc'] = config.CC_TO_ADDRESS
    msg['Bcc'] = ''
    all_recipients = [msg['To']] + msg['Cc'].split(',') + msg['Bcc'].split(',')
    msg['Subject'] = subject
    msg.attach(MIMEText(body_content, 'html'))
    # Add attachments if provided

    if file_name: 
        attachments = file_name
        print(attachments)
        for attachment in attachments:
            try:
                file_name = attachment.split('\\')[-1]
                with open(attachment, 'rb') as file:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(file.read())
                    encoders.encode_base64(part)  # Encode file in base64
                    part.add_header('Content-Disposition', f'attachment; filename={file_name}')  # Attachment filename
                    msg.attach(part)
                    logging.info(f"Attachment {attachment} added successfully.")
            except Exception as e:
                print(e)
                logging.error(f"Failed to attach {attachment}: {e}")
        
    try:
        smtp_server = config.SMTP_SERVER
        smtp_port = config.SMTP_PORT
        smtp_username = config.EMAIL_FROM
        smtp_password = config.SMTP_PASSWORD
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(smtp_username, smtp_password)
            logging.info(f"Mail sent to : {all_recipients}")
            server.sendmail(formatted_from, all_recipients, msg.as_string())
            print(f"Email sent to {all_recipients}")
            logging.info(f"Email sent to {all_recipients}")
            
    except Exception as e:
        print(f"Failed to send email to {to_address}: {e}")
        logging.info(f"Failed to send email to {to_address}: {e}")

def latest_employee_master_path(file_date):
    pattern = re.compile(r"Employee_Master_\d{2}_\d{2}_\d{4}\.xlsx$") 
    latest_file = os.path.join(config.EMPLOYEE_MASTER_PATH, rf"Employee_Master_{file_date}.xlsx")
    if os.path.isfile(latest_file) and os.path.exists(latest_file):
        path = latest_file
        return path
    else:
        list_of_files = [f for f in glob.glob(r"Z:\Harish\EM_Historical\Employee_Master_*.xlsx") if pattern.search(os.path.basename(f))]
        path = max(list_of_files, key=os.path.getctime)
        print(path)
        return path

# moving old code to backup folder
def move_to_yesterday_folder(REPORT_PATH):
    is_file_removed = False
    yesterday = (datetime.now() - timedelta(days = 1)).date()
    files = os.listdir(REPORT_PATH)
    valid_files = [os.path.join(REPORT_PATH, file) for file in files if  os.path.isfile(os.path.join(REPORT_PATH, file))]
    for file in valid_files:
        r_time = datetime.fromtimestamp(os.path.getmtime(file)).date()
        if r_time != datetime.now().date():
            is_file_removed = True 
            destination_path = os.path.join(REPORT_PATH, str(yesterday))
            os.makedirs(destination_path, exist_ok=True)
            shutil.move(file, os.path.join(destination_path))
    return is_file_removed

def sal_reconciliation():
    is_file_removed = move_to_yesterday_folder(REPORT_PATH)
    if is_file_removed:
        # download report
        call(['python', os.path.join(ROOT_DIR, 'salary_reconciliation_reports_download.py')])
    today = datetime.now().strftime('%d_%m_%Y')
    emp_bank_df = pd.read_excel(os.path.join(REPORT_PATH, 'PA0009.xlsx'))
    print("Emp_Bank loaded into memory.")
    emp_ctc_df = pd.read_excel(os.path.join(REPORT_PATH, 'PA0008.xlsx'))
    print("Emp_CTC loaded into memory.")
    emp_ctc_bankdescription_df = pd.read_excel(os.path.join(REPORT_PATH, 'BNKA.xlsx'))
    print("Emp_CTC_bankDesc loaded into memory.")
    emp_mol_df = pd.read_excel(os.path.join(REPORT_PATH, 'PA0185.xlsx'))
    print("Emp_MOL loaded into memory.")
    emp_org_df = pd.read_excel(os.path.join(REPORT_PATH, 'PA0001.xlsx'))
    print("Emp_ORG loaded into memory.")
    emp_status_df = pd.read_excel(os.path.join(REPORT_PATH, 'PA0000.xlsx'))
    print("Emp_Status loaded into memory.")
    employee_master = pd.read_excel(latest_employee_master_path(today))
    print("Emp_master loaded into memory.")

    # required columns
    emp_bank_df = emp_bank_df[['Personnel Number','End Date','Start Date','Bank Country/Region','Bank Key','Bank Account']]
    # mapp the bank (bankname) based on Bank Country/Region and Bank Key
    emp_ctc_bankdescription_df = emp_ctc_bankdescription_df[['Bank Country/Region','Bank Key','Name of Financial Institution']] 
    emp_ctc_df = emp_ctc_df[['Personnel Number','End Date','Start Date','Wage Type','Wage Type.1','Wage Type.2','Wage Type.3','Wage Type.4','Wage Type.5','Wage Type.6','Wage Type.7','Wage Type.8','Wage Type.9','Wage Type.10','Wage Type.11','Wage Type.12','Wage Type.13','Wage Type.14','Wage Type.15','Wage Type.16','Wage Type.17','Wage Type.18','Wage Type.19','Wage Type.20','Wage Type.21','Wage Type.22','Wage Type.23','Wage Type.24','Wage Type.25','Wage Type.26','Wage Type.27','Wage Type.28','Wage Type.29','Wage Type.30','Wage Type.31','Wage Type.32','Wage Type.33','Wage Type.34','Wage Type.35','Wage Type.36','Wage Type.37','Wage Type.38','Wage Type.39','Amount','Amount.1','Amount.2','Amount.3','Amount.4','Amount.5','Amount.6','Amount.7','Amount.8','Amount.9','Amount.10','Amount.11','Amount.12','Amount.13','Amount.14','Amount.15','Amount.16','Amount.17','Amount.18','Amount.19','Amount.20','Amount.21','Amount.22','Amount.23','Amount.24','Amount.25','Amount.26','Amount.27','Amount.28','Amount.29','Amount.30','Amount.31','Amount.32','Amount.33','Amount.34','Amount.35','Amount.36','Amount.37','Amount.38','Amount.39']]
    emp_ctc_df = emp_ctc_df.fillna(0.0)
    # converting the wages type to str
    emp_ctc_df[['Wage Type','Wage Type.1','Wage Type.2','Wage Type.3','Wage Type.4','Wage Type.5','Wage Type.6','Wage Type.7','Wage Type.8','Wage Type.9','Wage Type.10','Wage Type.11','Wage Type.12','Wage Type.13','Wage Type.14','Wage Type.15','Wage Type.16','Wage Type.17','Wage Type.18','Wage Type.19','Wage Type.20','Wage Type.21','Wage Type.22','Wage Type.23','Wage Type.24','Wage Type.25','Wage Type.26','Wage Type.27','Wage Type.28','Wage Type.29','Wage Type.30','Wage Type.31','Wage Type.32','Wage Type.33','Wage Type.34','Wage Type.35','Wage Type.36','Wage Type.37','Wage Type.38','Wage Type.39']] = emp_ctc_df[['Wage Type','Wage Type.1','Wage Type.2','Wage Type.3','Wage Type.4','Wage Type.5','Wage Type.6','Wage Type.7','Wage Type.8','Wage Type.9','Wage Type.10','Wage Type.11','Wage Type.12','Wage Type.13','Wage Type.14','Wage Type.15','Wage Type.16','Wage Type.17','Wage Type.18','Wage Type.19','Wage Type.20','Wage Type.21','Wage Type.22','Wage Type.23','Wage Type.24','Wage Type.25','Wage Type.26','Wage Type.27','Wage Type.28','Wage Type.29','Wage Type.30','Wage Type.31','Wage Type.32','Wage Type.33','Wage Type.34','Wage Type.35','Wage Type.36','Wage Type.37','Wage Type.38','Wage Type.39']].astype(int).astype(str)

    def calculate_conditional_total_amount(df: pd.DataFrame) -> pd.DataFrame:
        df['Total'] = 0.0
        column_pairs = []
        if 'Wage Type' in df.columns and 'Amount' in df.columns:
            column_pairs.append(('Wage Type', 'Amount'))

        for i in range(1, 40): # Iterate from 1 to 39
            wage_col_name = f'Wage Type.{i}'
            amount_col_name = f'Amount.{i}'
            if wage_col_name in df.columns and amount_col_name in df.columns:
                column_pairs.append((wage_col_name, amount_col_name))

        for wage_col, amount_col in column_pairs:
            wage_type_values = df[wage_col].astype(str).fillna('')
            amount_values = pd.to_numeric(df[amount_col], errors='coerce').fillna(0)
            df['Total'] += np.where(wage_type_values == '8500', 0, amount_values)
            
        return df

    emp_ctc_df = calculate_conditional_total_amount(emp_ctc_df.copy())
    print(emp_ctc_df)
    # MoL
    emp_mol_df = emp_mol_df[['Personnel Number','End Date','Start Date','Identity Number']]
    # client code 
    emp_org_df = emp_org_df[['Personnel Number','End Date','Start Date','Organizational unit']]
    # Employee Status - TA and 3
    emp_status_df = emp_status_df[['Personnel Number','End Date','Start Date','Action Type','Employment status']]

    employee_master = employee_master[['Employee Code/Applicant ID','Full Name','Employee Status','Client Code','Total','Air Ticket','accountNumber','bank (bankName)','MOL Number']]
    emp_ctc_df = emp_ctc_df[['Personnel Number','End Date','Start Date','Total']]

    # employee_master['Air Ticket'] = employee_master['Air Ticket'].fillna(0.0)
    employee_master['EM_Total'] = employee_master['Total'] 
    employee_master.drop(columns={'Total'}, inplace=True)
    # Sort all DataFrames by 'Start Date' in descending order
    emp_bank_df = emp_bank_df.sort_values(by='Start Date', ascending=False)
    emp_ctc_bankdescription_df = emp_ctc_bankdescription_df.sort_values(by='Bank Country/Region', ascending=True)  # key reference table — no Start Date
    emp_ctc_df = emp_ctc_df.sort_values(by='Start Date', ascending=False)
    emp_mol_df = emp_mol_df.sort_values(by='Start Date', ascending=False)
    emp_org_df = emp_org_df.sort_values(by='Start Date', ascending=False)
    emp_status_df = emp_status_df.sort_values(by='Start Date', ascending=False)
    employee_master = employee_master.sort_values(by='Employee Code/Applicant ID', ascending=True)  # static master data
    employee_master = employee_master[employee_master['Employee Status']=='Active'] # Active Emps
    # Keep latest row per Personnel Number (after sort)
    emp_bank_df = emp_bank_df.drop_duplicates(subset='Personnel Number', keep='first')
    emp_ctc_df = emp_ctc_df.drop_duplicates(subset='Personnel Number', keep='first')
    emp_mol_df = emp_mol_df.drop_duplicates(subset='Personnel Number', keep='first')
    emp_org_df = emp_org_df.drop_duplicates(subset='Personnel Number', keep='first')
    emp_status_df = emp_status_df.drop_duplicates(subset='Personnel Number', keep='first')
    emp_bank_df = emp_bank_df.fillna('')
    emp_ctc_df = emp_ctc_df.fillna('')
    emp_mol_df = emp_mol_df.fillna('')
    emp_org_df = emp_org_df.fillna('')
    emp_status_df = emp_status_df.fillna('')
    employee_master = employee_master.fillna('')  
    emp_ctc_df['Total'] = emp_ctc_df['Total'].astype(int)
    emp_status_df['Employment status'] = emp_status_df['Employment status'].apply(lambda x: 'Active' if x == 3 else 'Inactive')
    employee_master[['Client Code','EM_Total']] = employee_master[['Client Code','EM_Total']].astype(int)
    emp_bank_df = emp_bank_df.rename(columns={
    'Start Date': 'Start Date Bank',
    'End Date': 'End Date Bank'
    })

    emp_ctc_df = emp_ctc_df.rename(columns={
        'Start Date': 'Start Date CTC',
        'End Date': 'End Date CTC'
    })

    emp_status_df = emp_status_df.rename(columns={
        'Start Date': 'Start Date Status',
        'End Date': 'End Date Status'
    })

    emp_mol_df = emp_mol_df.rename(columns={
        'Start Date': 'Start Date MoL',
        'End Date': 'End Date MoL'
    })

    emp_org_df = emp_org_df.rename(columns={
        'Start Date': 'Start Date Org',
        'End Date': 'End Date Org'
    })

    # Merge emp_bank_df with emp_ctc_bankdescription_df on Country and Bank Key
    emp_bank_df = emp_bank_df.merge(
        emp_ctc_bankdescription_df,
        how='left',
        on=['Bank Country/Region', 'Bank Key']
    )
    # taking only the Active employees 
    active_employees_df = emp_status_df[(emp_status_df['Action Type'] != 'TF')].copy()
    initial_active_count = len(active_employees_df) 
    print(f"Total active employees identified from emp_status_df: {initial_active_count}")
    # Converting personnel number as str 
    active_employees_df['Personnel Number'] = active_employees_df['Personnel Number'].astype(str)
    emp_bank_df['Personnel Number'] = emp_bank_df['Personnel Number'].astype(str)
    emp_ctc_df['Personnel Number'] = emp_ctc_df['Personnel Number'].astype(str)
    emp_org_df['Personnel Number'] = emp_org_df['Personnel Number'].astype(str)
    emp_mol_df['Personnel Number'] = emp_mol_df['Personnel Number'].astype(str) 

    active_p_numbers = active_employees_df['Personnel Number'].unique() # Get a set for faster lookup
    emp_bank_df = emp_bank_df[emp_bank_df['Personnel Number'].isin(active_p_numbers)].copy() #
    emp_ctc_df = emp_ctc_df[emp_ctc_df['Personnel Number'].isin(active_p_numbers)].copy()
    emp_org_df = emp_org_df[emp_org_df['Personnel Number'].isin(active_p_numbers)].copy()
    emp_mol_df = emp_mol_df[emp_mol_df['Personnel Number'].isin(active_p_numbers)].copy()

    merged_df = active_employees_df.merge(emp_bank_df, on='Personnel Number', how='inner')
    merged_df = merged_df.merge(emp_ctc_df, on='Personnel Number', how='inner')
    merged_df = merged_df.merge(emp_org_df, on='Personnel Number', how='inner')
    # Split based on Personnel Number series
    merged_non_6 = merged_df[~merged_df['Personnel Number'].astype(str).str.startswith('6')]
    merged_6 = merged_df[merged_df['Personnel Number'].astype(str).str.startswith('6')]

    # Merge only the non-6 employees with MoL data
    merged_non_6 = merged_non_6.merge(emp_mol_df, on='Personnel Number', how='inner')

    # Recombine both sets
    final_merged_df = pd.concat([merged_non_6, merged_6], ignore_index=True)
    employee_master['Employee Code/Applicant ID'] = employee_master['Employee Code/Applicant ID'].astype(str)
    final_df = final_merged_df.merge(
    employee_master,
    how='outer',
    left_on='Personnel Number',
    right_on='Employee Code/Applicant ID'
    )
    final_df = final_df.fillna('')
    final_df['Client Code'] = final_df['Client Code'].replace('',0).astype(int)
    final_df['Organizational unit'] = final_df['Organizational unit'].replace('',0).astype(int)
    final_df['Total'] = final_df['Total'].replace('',0).astype(int)
    final_df['EM_Total'] =  final_df['EM_Total'].replace('',0).astype(int)
    bank_bankname_processed = final_df['bank (bankName)'].astype(str).str.lower().str.replace(' ', '', regex=False).fillna('')
    name_fin_inst_processed = final_df['Name of Financial Institution'].astype(str).str.lower().str.replace(' ', '', regex=False).fillna('')

    # Find mismatches by checking which employee_master fields don't match merged_df
    # Create validation columns (True if matches, False if not)
    
    final_df['Client Code match'] = final_df['Client Code'] == final_df['Organizational unit']
    final_df['Employee Status match'] = final_df['Employee Status'] == final_df['Employment status']
    final_df['Bank Name match'] = bank_bankname_processed == name_fin_inst_processed
    final_df['MOL Number match'] = final_df['MOL Number'] == final_df['Identity Number']
    final_df['Total match'] = final_df['Total'] == final_df['EM_Total']

    reconciliation_issues = final_df[
        (final_df['Client Code'] != final_df['Organizational unit']) |
        (final_df['Employee Status'] != final_df['Employment status']) |
        (bank_bankname_processed != name_fin_inst_processed) | 
        (final_df['MOL Number'] != final_df['Identity Number']) |
        (final_df['Total'] != final_df['EM_Total'])
    ]
    reconciliation_issues = reconciliation_issues[['Personnel Number','Client Code','Organizational unit','Employee Status','Employment status','bank (bankName)','Name of Financial Institution', 'MOL Number','Identity Number','Total','EM_Total']]
    output_filepath = os.path.join(OUTPUT_PATH, f"mismatch_employee_details_withoutBank_{today}.xlsx")
    reconciliation_issues.to_excel(output_filepath, index=False)
    final_df = final_df[['Personnel Number','Employee Code/Applicant ID', 'Full Name','Action Type','accountNumber','Bank Country/Region', 'Bank Key','Air Ticket','Bank Account', 'Client Code','Organizational unit', 'Client Code match','Employment status', 'Employee Status', 'Employee Status match','bank (bankName)','Name of Financial Institution','Bank Name match','MOL Number', 'Identity Number','MOL Number match','Total', 'EM_Total', 'Total match']]
    final_df.to_excel(output_filepath.replace('.xlsx','_Overall.xlsx'),index=False)
    # send mail 
    to_address = config.SALRECON_TO_ADDRESS
    subject = f'Salary Reconcilation_{today}'
    body_content = "<html><body>Hi Team, <br><br>Today's Salary reconcilation attachment attached in the mail, please check.<br><br><br><br> Warm Regards, <br><br> TASC Group.<br></body></html>"
    file_name = [output_filepath, output_filepath.replace('.xlsx','_Overall.xlsx')]
    send_email(to_address, subject, body_content, file_name)
    print('success')

if __name__ == '__main__':
    sal_reconciliation()

# How to run this script
# C:\Users\BotServer\tascautomations\tasc_venv\Scripts\python.exe C:\Users\BotServer\tascautomations\salary_reconcilation\salary_reconciliation.py