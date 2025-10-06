import pandas as pd
from datetime import datetime, timedelta
import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage  # to load image while sending in email
import smtplib


class Validator:
    def __init__(self):
        self.base_file_path = r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\Early_Warning_supporting_Reports\nandhini\baseReport.xlsx'
        self.employee_status_file_path_pa0000 = r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\Early_Warning_supporting_Reports\nandhini\PA0000.xlsx'
        self.payroll_status_file_path_pa0001 = r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\Early_Warning_supporting_Reports\nandhini\PA0001.xlsx'
        self.salary_file_path_pa0008 = r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\Early_Warning_supporting_Reports\nandhini\PA0008.xlsx'
        self.mol_file_path_pa0185 = r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\Early_Warning_supporting_Reports\nandhini\PA0185.xlsx'
        self.bnk_ac_no_file_path_pa0009 = r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\Early_Warning_supporting_Reports\nandhini\PA0009.xlsx'
        self.nationality_file_path_pa0002 = r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\Early_Warning_supporting_Reports\nandhini\PA0002.xlsx'
        self.nationality_mapping = r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\Early_Warning_supporting_Reports\nandhini\Nationality_Mapping.xlsx'
        self.kna_file_path = r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\Early_Warning_supporting_Reports\nandhini\KNA1.xlsx'
    
    def read_excel_file(self):
        
        self.base_df = pd.read_excel(self.base_file_path).fillna('')
        self.employee_status_df = pd.read_excel(self.employee_status_file_path_pa0000).fillna('')
        self.payroll_status_df = pd.read_excel(self.payroll_status_file_path_pa0001).fillna('')
        self.salary_status_df = pd.read_excel(self.salary_file_path_pa0008).fillna('')
        self.mol_status_df = pd.read_excel(self.mol_file_path_pa0185).fillna('')
        self.bnk_ac_no_status_df = pd.read_excel(self.bnk_ac_no_file_path_pa0009).fillna('')
        self.nationality_status_df = pd.read_excel(self.nationality_file_path_pa0002).fillna('')
        self.kna_df = pd.read_excel(self.kna_file_path).fillna('')

    def mark_duplicate(self):
        check_columns = ['iban', 'MOL Number', 'National ID', 'Iquama Number']
        for col in check_columns:
            # Find all duplicate groups (ignoring NaN / blank)
            dup_groups = (
                self.base_df[self.base_df[col].notna()]
                .groupby(col)['Employee Code/Applicant ID']
                .apply(list)           # collect all Employee Codes with same value
                .reset_index(name='EmpCodes')
            )
            
            # Filter groups with actual duplicates (more than 1 code)
            dup_groups = dup_groups[
                (dup_groups['EmpCodes'].str.len() > 1) & 
                (dup_groups[col].notna()) & 
                (dup_groups[col].astype(str).str.strip() != '')
            ]

            
            # Loop through each duplicate group and update status
            for _, row in dup_groups.iterrows():
                mol_value = row['MOL Number']
                emp_codes = row['EmpCodes']

                # Loop over each employee in the group
                for i, emp in enumerate(emp_codes):
                    # Pick the next employee in the list to point to
                    # If it's the last one, point to the first one to keep it consistent
                    target_emp = emp_codes[(i + 1) % len(emp_codes)]

                    dup_text = f"{col} Duplicate with {target_emp}"

                    self.final_df.loc[
                        self.final_df['Employee Code/Applicant ID'] == emp,
                        'Status'
                    ] += f"{dup_text}, "

    def send_typist_email(self, sender_email, sender_password, recipient_email, body, data, attachment_dir=None):

        msg_root = MIMEMultipart('related')
        msg_root['From'] = sender_email
        msg_root['To'] = recipient_email
        msg_root['Cc'] = ''
        msg_root['Bcc'] = ''
        msg_root['Subject'] = data['subject']

        all_recipients = [msg_root['To']] + msg_root['Cc'].split(',') + msg_root['Bcc'].split(',')

        # HTML body part
        msg_alternative = MIMEMultipart('alternative')
        msg_root.attach(msg_alternative)

        # Optional attachments
        if attachment_dir:
            for attachment in attachment_dir:
                try:
                    file_name = attachment.split('\\')[-1]
                    with open(attachment, 'rb') as file:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(file.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename={file_name}')
                        msg_root.attach(part)
                        print(f"Attachment {attachment} added successfully.")
                except Exception as e:
                
                    print(f"Failed to attach {attachment}: {e}")

        # Send mail
        try:
            smtp_server = "smtp.office365.com"
            smtp_port = 587
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, all_recipients, msg_root.as_string())
                print(f"Email sent to {all_recipients}")
        except Exception as e:
            print(f"Failed to send email to {recipient_email}: {e}")
        

    def get_field_mapping(self):
        self.base_df = self.base_df[self.base_df['Employee Status'] == 'Active']
        # self.base_df = self.base_df[0:10]
        content = []
        for index,row in self.base_df.iterrows():
            fields = {'Status': ''}
            fields.update(row.to_dict())
            mol_status_record = self.mol_status_df[self.mol_status_df['Personnel Number'] == row['Employee Code/Applicant ID']]
            employee_status_record = self.employee_status_df[self.employee_status_df['Personnel Number'] == row['Employee Code/Applicant ID']]
            salary_status_record = self.salary_status_df[self.salary_status_df['Personnel Number'] == row['Employee Code/Applicant ID']]
            bnk_ac_no_status_record = self.bnk_ac_no_status_df[self.bnk_ac_no_status_df['Personnel Number'] == row['Employee Code/Applicant ID']]
            nationality_status_record = self.nationality_status_df[self.nationality_status_df['Personnel Number'] == row['Employee Code/Applicant ID']]
            payroll_status_record = self.payroll_status_df[self.payroll_status_df['Personnel Number'] == row['Employee Code/Applicant ID']]
            kna_status_record = self.kna_df[self.kna_df['Customer'] == row['Employee Code/Applicant ID']]
            
            #Employee Status
            employee_status = employee_status_record.sort_values(by='Start Date', ascending=False)[['Start Date','Personnel Number','Employment status']]
            employee_status['Employment status'] = employee_status['Employment status'].apply(lambda x: 'Active' if x == 3 else 'Inactive')
            if employee_status.empty:
                fields['Status'] += 'Employee Missing,'
            
            else:
                if employee_status.head(1)['Employment status'].iloc[0] == '':
                    fields['Status'] += 'Employee Missing,'
                if employee_status.head(1)['Employment status'].iloc[0] != row['Employee Status']:
                    fields['Status'] += 'Employee Status Mismatch, '
                # if employee_status.head(1)['Start Date'].iloc[0] != row['Employment Details Hire Date']:
                #     fields['Status'] += 'Employment Details Hire Date Mismatch, '

            #MOL Number
            mol_status = mol_status_record.sort_values(by='Start Date', ascending=False)[['Start Date','Personnel Number','Identity Number', 'Expiry date']]
            mol_status['Expiry date'] = mol_status['Expiry date'].astype(str).replace('NaT', '')

            if mol_status.empty:
                fields['Status'] += 'MOL Number Missing, '
            else:
                if mol_status.head(1)['Identity Number'].iloc[0] != row['MOL Number']:
                    fields['Status'] += 'MOL Number Mismatch, '
                if len(mol_status.head(1)['Identity Number'].iloc[0]) != 14:
                    fields['Status'] += 'MOL Number Length Issue'
                if mol_status.head(1)['Identity Number'].iloc[0] == '':
                    fields['Status'] += 'MOL Number Missing in SF, '
                #MOL Expiry Date
                if mol_status.head(1)['Expiry date'].iloc[0] == '':
                    fields['Status'] += 'MOL Expiry Date Missing, '
                if mol_status.head(1)['Expiry date'].iloc[0] != row['MOL Expiry Date'].strftime('%Y-%m-%d'):
                    fields['Status'] += 'MOL Expiry Date Mismatch, '
                cur_date = datetime.today()
                expiry_date = row['MOL Expiry Date']
                days_diff = (expiry_date - cur_date).days

                if days_diff < -60:
                    fields['Status'] = "❌ Error: MOL expired more than 60 days ago, "
                elif days_diff <= 0:
                    fields['Status'] = "⚠ Warning: MOL expired within last 60 days, "

            #Account Number
            bnk_status = bnk_ac_no_status_record.sort_values(by='Start Date', ascending=False)[['Start Date','Personnel Number','Bank Account', 'IBAN', 'Payment Method']]
            if bnk_status.empty:
                fields['Status'] += 'Bank Details Missing, '
            else:

                # --- Bank Account Checks ---
                # if bnk_status.head(1)['Bank Account'].iloc[0] != row['accountNumber']:
                #     fields['Status'] += 'Bank Account Number Mismatch, '

                # --- IBAN Checks ---
                if bnk_status.head(1)['IBAN'].iloc[0] == '':
                    fields['Status'] += 'IBAN Number Missing, '
                if bnk_status.head(1)['IBAN'].iloc[0] != row['iban']:
                    fields['Status'] += 'IBAN Number Mismatch, '
                if len(bnk_status.head(1)['IBAN'].iloc[0]) != 24:
                    fields['Status'] += 'IBAN Length Issue'
                

                # if bnk_status['Payment Method'].iloc[0] != row['paymentMethod (Label)']:
                #     fields['Status'] += 'Payment Method Mismatch, '
                    
            #Salary
            salary_status = salary_status_record.sort_values(by='Start Date', ascending=False)
            if salary_status.empty:
                fields['Status'] += 'Salary Details Missing, '
            else:
                cols_with_amount = [col for col in salary_status.columns if 'amount' in col.lower()]
                salary_status = salary_status[['Start Date','Personnel Number']+cols_with_amount]
                if salary_status.head(1)[cols_with_amount].sum(axis=1).iloc[0] == 0:
                    fields['Status'] += 'Salary Amount Missing, '
                if salary_status.head(1)[cols_with_amount].sum(axis=1).iloc[0] != row['Total']:
                    fields['Status'] += 'Salary Amount Mismatch, '

            #Nationality
            # nationality_status = nationality_status_record.sort_values(by='Start Date', ascending=False)
            # if nationality_status.empty:
            #     fields['Status'] += 'Nationality Details Missing, '
            # else:
            #     nationality_map = pd.read_excel(self.nationality_mapping)
            #     nationality_map = nationality_map[nationality_map['Country/Region Key'] == nationality_status['Nationality'].iloc[0]]
            #     if nationality_map['Country/Region Name'].iloc[0] == '':
            #         fields['Status'] += 'Nationality Missing'
            #     if nationality_map['Country/Region Name'].iloc[0] != row['Nationality']:
            #         fields['Status'] += 'Nationality Mismatch, '
            #     if nationality_status['Date of birth'].iloc[0] == '':
            #         fields['Status'] += 'Date of birth Missing'
            #     if nationality_status['Date of birth'].iloc[0] != row['Date Of Birth']:
            #         fields['Status'] += 'Date of birth Mismatch, '

            #Payroll
            payroll_status = payroll_status_record.sort_values(by='Start Date', ascending=False)
            if payroll_status.empty:
                fields['Status'] += 'Payroll Details Missing, '
            else:
                if payroll_status.head(1)['Organizational unit'].iloc[0] == '':
                    fields['Status'] += 'Payroll Details Missing'
                if payroll_status.head(1)['Organizational unit'].iloc[0] != row['Client Code']:
                    fields['Status'] += 'Client Code Mismatch, '
                if row['Pay Group']:
                    if payroll_status.head(1)['Payroll area'].iloc[0] == '':
                        fields['Status'] += 'Payroll area Missing'
                    if payroll_status.head(1)['Payroll area'].iloc[0] != row['Pay Group'].split('-')[0]:
                        fields['Status'] += 'Payroll Area Mismatch, '
            print(fields)
            content.append(fields)
        self.final_df = pd.DataFrame(content)
        # self.mark_duplicate('Bank Account', 'accountNumber', final_df)
        self.mark_duplicate()
        
        highlight_color = 'background-color: #FFBF00'   # shade close to your screenshot

        # Create a styled object
        styled_df = self.final_df.style.apply(
            lambda col: [highlight_color if str(x).strip() != '' else '' for x in col],
            subset=['Status']
        )

        # Export with styles applied
        styled_df.to_excel(
            r"C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\Early_Warning_supporting_Reports\validation.xlsx",
            index=False
        )
        print('Success')
        
#['Employee Code/Applicant ID', 'Employee Status', 'Employment Details Hire Date', 'MOL Number', 'MOL Expiry Date', 'accountNumber', 'iban', 'paymentMethod (Label)', 'Total', 'Nationality', 'Date Of Birth', 'Client Code', 'Pay Group']
obj = Validator()
obj.read_excel_file()
obj.get_field_mapping()
