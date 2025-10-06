import os
import requests
import pandas as pd
import smtplib
from bs4 import BeautifulSoup
from email.message import EmailMessage
from datetime import datetime
from dateutil.parser import parse

class AutoMail:

    def __init__(self):
        self.today = datetime.now().strftime('%d-%b-%Y')
        self.month = datetime.now().month
        self.year = datetime.now().year
        self.month = 4
        self.year = 2024
    def send_email(self, sender_email, sender_password, recipient_email, subject, body):

        msg = EmailMessage()
        msg["From"] = sender_email
        msg["To"] = recipient_email
        msg["Subject"] = subject
        msg.set_content(body)
        msg.add_alternative(body, subtype='html')

        try:
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                server.login(sender_email, sender_password)
                server.send_message(msg)
            print("Email sent successfully.")

        except Exception as e:
            print(f"Failed to send email: {e}")
    
    def format_body_content(self, data):
        content = data.to_html(index=False)
        msg_body = '''Hi Team, <br><br> Please find the below details:<br><br>'''
        msg_body += '''<html>
        <head>
        <style>
            table {
            border-collapse: collapse;
            width: 80%;
            font-family: Arial, sans-serif;
            text-align: center;
            }

            th {
            background-color: #003366;
            color: white;
            font-weight: bold;
            border: 1px solid black;
            padding: 8px;
            text-align: center;
            }

            td {
            border: 1px solid black;
            padding: 8px;
            
            }
        </style>
        </head>'''+ content
        msg_body += '''<br><br>Note: This is auto generated email please do not reply.'''

        return msg_body
    
    def get_pending_employee_email_list(self, code):
        url = f"https://api22preview.sapsf.com/odata/v2/PerEmail?$filter=personIdExternal eq '{code.strip()}'"
        payload = {}
        headers = {}
        response = requests.request("GET", url, headers=headers, data=payload, auth=('SFAPI@tasclabourT1', 'Exalogic@123456'))
        soup = BeautifulSoup(response.text, 'xml')
        content = [i for i in soup.find_all('properties') if i.find('emailType').text == '3204']
        email_list = [i.find('emailAddress').text  for i in content]
        # email_list = ['nandhini2571999@gmail.com','nandylokesh123@gmail.com']
        return email_list

    def timesheet_data(self):
        '''EmployeeCode, EmployeeNAme, Timesheet submission Date, Current status, pending with whome'''
        url = f"https://tascdev-interested-mouse-bv.cfapps.eu10-004.hana.ondemand.com/api/ts/V1/getEmployeeTimesheetSummary?month={self.month}&year={self.year}"

        payload = {}
        headers = {'security-key': '654321'}

        response = requests.request("GET", url, headers=headers, data=payload)
        if response.status_code == 200:
            json_data = response.json()
            data = json_data['data']
            df = pd.DataFrame(data)
            df = df.fillna('')
            df_filter = df[df['final_status'] == 'Pending']
            df_filter['timesheet_submission'] = pd.to_datetime(df_filter['timesheet_submission']).dt.strftime('%d-%b-%Y')
            pending_lm_code_list = list(set(df_filter['pending_employee_code'].tolist()))            
            for code in pending_lm_code_list:
                email_list = self.get_pending_employee_email_list(code)
                content = df_filter[df_filter['pending_employee_code'] == code][['employee_id','employee_name', 'timesheet_submission', 'current_status', 'pending_employee_code']]
                content = content.rename(columns = {'employee_id':'EmployeeCode','employee_name':'EmployeeName', 'timesheet_submission':'TimeSheet Submission Date', 'current_status': 'Current Status', 'pending_employee_code': 'Pending Employee Code'})
                sender_email = "nandhini2571999@gmail.com" 
                sender_password = "utfy ogzm qwks vfwt"  
                recipient_email = ','.join(email_list)
                subject = f"Pending Employee TimeSheet Alert!!! - {self.today}"
                body = self.format_body_content(content)
                self.send_email(sender_email, sender_password, recipient_email, subject,body)
                print("Email Sent Successfully!")

obj = AutoMail()
timesheet_info = obj.timesheet_data()


