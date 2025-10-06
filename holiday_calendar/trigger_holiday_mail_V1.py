import os
import pandas as pd
import smtplib
from email.message import EmailMessage
from datetime import datetime
from dateutil.parser import parse
class Holiday:

    def __init__(self):
        self.holiday_file = r"C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\holiday_calendar\Holiday Calendar for 2025 1 (1).xlsx"

    def send_holiday_email(self, sender_email, sender_password, recipient_email, subject, body):

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
    
    def formated_date(self, row):
        format_date = '1900-01-01'
        try:
            if ':' in str(row['Date']):
                format_date = datetime.strptime(str(row['Date']), "%Y-%m-%d %H:%M:%S").strftime("%Y-%m-%d")
            else:
                format_date = datetime.strptime(str(row['Date']), "%d-%m-%Y").strftime("%Y-%m-%d")
        except:
            format_date = row['Date'] 
        return format_date
                  
    def monthly_holiday_data(self):
        sheets_dict = pd.read_excel(self.holiday_file, sheet_name=None)
        content = []
        for sheet_name, sheet_df in sheets_dict.items():
            df = sheet_df.fillna('')
            found = False

            for index, row in df.iterrows():
                format_date = self.formated_date(row)
                if format_date:
                    # Check if month matches result_date
                    if result_date.month == int(str(format_date).split('-')[1]):
                        found = True
                        content.append({
                            'Date': format_date,
                            'Holiday': row.get('Holiday', 'Name'),
                            'Region': sheet_name
                        })

            # If no holiday found for that region
            if not found:
                content.append({
                    'Date': '',
                    'Holiday': f"No official holiday scheduled for {result_date.strftime('%B')} Month",
                    'Region': sheet_name
                })

        unique_data = [dict(t) for t in {tuple(d.items()) for d in content}]
        unique_data = sorted(unique_data, key = lambda x: x['Region'])
        sorted_data = sorted(unique_data,    key=lambda x: (x['Date'] == '', x['Date']))
        msg_body = f'''
        Dear All, \n<br><br>
        
        Please find the list of Holiday expected in month of {result_date.strftime('%B')} <br><br>

                    <table border="1" style="border-collapse: collapse;">
                <tr>
                    <th>Region</th>
                    <th>Date</th>
                    <th>Holiday</th>
                </tr>'''
        for i in sorted_data:
            if f"No official holiday scheduled for {result_date.strftime('%B')} Month" not in i['Holiday']:
                msg_body += f'''
                <tr>
                    <td>{i['Region']}</td>
                    <td>Holiday Confirmed on {i['Date']} for both public and private sectors</td>
                    <td>{i['Holiday']}</td>
                </tr>'''
            else:
                msg_body += f'''
                <tr>
                    <td>{i['Region']}</td>
                    <td align="center">-</td>
                    <td>{i['Holiday']}</td>
                </tr>'''
        msg_body += '''</table><br><br>'''
        msg_body += '''Note: This is TASC Auto generated email. Please don't respond.<br><br> Thanks and Regards,<br>TASC Automation'''
        return msg_body


today = datetime.now()

if today.day < 15:
    result_date = today.replace(day=5)
else:
    # Handle month rollover
    if today.month == 12:
        result_date = today.replace(year=today.year + 1, month=1, day=5)
    else:
        result_date = today.replace(month=today.month + 1, day=5)

print(result_date.strftime("%Y-%m-%d")) 

obj = Holiday()
msg_body = obj.monthly_holiday_data()
sender_email = "nandhini2571999@gmail.com"  
sender_password = "utfy ogzm qwks vfwt"  
recipient_email = "nandylokesh123@gmail.com" 
subject = f"Holiday List for {result_date.strftime('%B')} – GCC Region"
body = msg_body

obj.send_holiday_email(sender_email, sender_password, recipient_email, subject,body)