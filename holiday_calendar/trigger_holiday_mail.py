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

    def monthly_holiday_data(self):
        sheets_dict = pd.read_excel(self.holiday_file, sheet_name=None)
        content = []
        for sheet_name in sheets_dict.keys():
            df = sheets_dict[sheet_name].fillna('')
            for index,row in df.iterrows():
                format_date = '1900-01-01'
                try:
                    format_date = parse(str(row['Date'])).strftime("%Y-%m-%d")
                except:
                    format_date = row['Date']                     

                if format_date and datetime.now().month == int(str(format_date).split('-')[1]):
                    fields = {'Date': format_date,
                                'Holiday': row.get('Holiday','Name'),
                                'Region': sheet_name}
                    content.append(fields)

                if format_date == '':
                    fields = {'Date': '',
                                'Holiday': 'No official holiday announced',
                                'Region': sheet_name}
                    content.append(fields)
                
                msg_body = '''
                Dear All, \n
                
                Please find below the update regarding the Islamic New Year holiday in various countries: 

                            <table border="1" style="border-collapse: collapse;">
                        <tr>
                            <th>Region</th>
                            <th>Holiday</th>
                        </tr>'''
                for i in content:
                    if 'No official holiday announced' not in i['Holiday']:
                        msg_body += f'''
                        <tr>
                            <td align="center">{i['Region']}</td>
                            <td>Holiday Confirmed on {i['Date']} for both public and private sectors</td>
                        </tr>'''
                    else:
                        msg_body += f'''
                        <tr>
                            <td align="center">{i['Region']}</td>
                            <td>No official holiday announced</td>
                        </tr>'''
                msg_body += '''</table>'''
        return msg_body


today = datetime.now().strftime('%d-%b-%Y')
obj = Holiday()
msg_body = obj.monthly_holiday_data()
sender_email = "nandhini2571999@gmail.com"  
sender_password = "utfy ogzm qwks vfwt"  
recipient_email = "nandylokesh123@gmail.com" 
subject = f"Holiday Calendar - {today}"
body = msg_body

obj.send_holiday_email(sender_email, sender_password, recipient_email, subject,body)