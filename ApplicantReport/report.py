import pandas as pd
import smtplib
from email.message import EmailMessage

class AppNotification:
    def __init__(self):
        self.path = r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\ApplicantReport\ProVisaDetailedReport.csv'
        self.today_date = '2025-09-25'#pd.to_datetime('today').strftime('%Y-%m-%d')

    def send_gmail_email(self, sender_email, sender_password, recipient_email, subject, body):

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

    def send_email(self, sender_email, sender_password, recipient_email, subject, body):
        self.email_status = False
        msg = EmailMessage()
        msg["From"] = sender_email
        msg["To"] = recipient_email
        msg["Subject"] = subject
        msg.set_content(body)
    
        try:
            
            # Send email using Outlook SMTP
            with smtplib.SMTP("smtp.office365.com", 587) as server:
                server.starttls()  # Use TLS (not SSL)
                server.login(sender_email, sender_password)
                server.send_message(msg)
                self.email_status = True
            print("Email sent successfully via Outlook.")

        except Exception as e:
            print(f"Failed to send email: {e}")

    def get_residence_mail_notification(self, df):

        df_filter = df[df['Visa Request Type'].isin(['New 1 YR','New']) & df['Visa Type'].isin(['Residence Visa','Residence Visa DHCC','Residence Visa TECOM']) & df['Activity Name'].isin(['Approval of Residence Stamping','Process Completed'])]

        df_filter_final = df_filter[df_filter['Activity End Date'] == self.today_date]
        df_filter_final = df_filter_final[0:1]
        for index, row in df_filter_final.iterrows():
            subject = f"Test Residency Visa approval Completed_{row['Applicant Name']}_{row['Client Name']}_ {row['Application ID']}"
            html_body = f"""
            <html>
            <body style="font-family: Arial, sans-serif; font-size:14px; line-height:1.6;">
                <p>Dear <b>{row['Applicant Name']}</b>,</p>

                <p>Congratulations! Your <b>Residence Visa</b> has been approved, and the process has been successfully completed.</p>

                <p>Please note that it is no longer necessary to provide your original passport to complete the Visa stamping process. 
                You will receive a notification once your <b>e-Visa stamping</b> is completed, after which you can easily download your 
                stamped e-Visa via the UAE ICP app.</p>

                <h4>How to Download e-Residence and Emirates ID Copies from the ICP App:</h4>
                <ol>
                <li>Download the ICP mobile application (Available in 
                    <a href="https://play.google.com/store">Play Store</a> and 
                    <a href="https://www.apple.com/app-store/">Apple Store</a>).
                </li>
                <li>Login using your Emirates ID number.</li>
                <li>Download the e-Residence copy and Emirates ID copy.</li>
                </ol>

                <h4>How to collect your Emirates ID card:</h4>
                <p>After your Emirates ID card is printed, the UAE Authority will send you a notification via SMS to your registered 
                mobile number. Once you receive this message, there are two options available for collecting your Emirates ID card:</p>

                <ol>
                <li><b>Direct to Your Location:</b> You can choose the option to have the Emirates ID sent to your location.
                    Once you receive the Emirates ID card, we request you to share a scanned copy of the card 
                    (both front and back sides) with us at 
                    <a href="mailto:documents@tascoutsourcing.com">documents@tascoutsourcing.com</a>.
                </li>

                <li><b>Through TASC:</b> Please share the SMS you receive from the UAE Authority with us at 
                    <a href="mailto:joiningauh@tascoutsourcing.com">joiningauh@tascoutsourcing.com</a>, 
                    so our PRO can collect your Emirates ID from the centre. 
                    You can then collect the Emirates ID from the TASC office.
                </li>
                </ol>

                <p>If you have any further questions, please do not hesitate to reach out to us.</p>

                <p>We value your feedback, so please share your thoughts with us via our feedback survey at 
                <a href="mailto:letstalk@tascoutsourcing.com">letstalk@tascoutsourcing.com</a> or call 800-ECARE (80032273) 
                from Monday to Friday, 8:00 am to 5:30 pm.</p>

                <p>Regards,<br>
                <b>TASC Team</b>
                </p>
            </body>
            </html>
            """
            sender_email = "nandhini2571999@gmail.com"  
            sender_password = "utfy ogzm qwks vfwt"  
            recipient_email = ["arti@tascoutsourcing.com","basha@tascoutsourcing.com","kokila.v@tascoutsourcing.com"]
            self.send_gmail_email(sender_email, sender_password, recipient_email, subject, html_body)
        
    def get_mol_mail_notification(self, df):
        
        df_filter = df[df['Visa Request Type'].isin(['New']) & df['Visa Type'].isin(['Residence Visa']) & df['Company'].isin(['TASC Labour Services LLC','Top Talent Employment services LLC']) & df['Activity Name'].isin(['Receipt of MOL Approval'])]

        df_filter_final = df_filter[df_filter['Activity End Date'] == self.today_date]
        df_filter_final = df_filter_final[0:1]
        for index,row in df_filter_final.iterrows():
            subject = f"Test MOL Approval Received {row['Applicant Name']}_{row['Client Name']}_{row['Application ID']}"
            html_body = f"""
            <html>
            <body style="font-family: Arial, sans-serif; font-size:14px; line-height:1.5;">
                <p>Dear <b>{row['Applicant Name']}</b>,</p>

                <p>Congratulations! We have received your MOHRE (Ministry of Human Resource) approval for employment with <b>TASC Group</b>.</p>

                <p>Here are the key milestones of your onboarding process at TASC Group:</p>

                <img src="cid:residence_process_image" alt="Residence Approval Process" style="max-width:300px;"><br><br>

                <p>Should you have any further questions, please do not hesitate to raise a ticket on our Chat bot – 
                <b>AIDA +97143588500</b> or write to us at 
                <a href="mailto:joining@tascoutsourcing.com">joining@tascoutsourcing.com</a>.
                </p>

                <p>We will be available during the weekdays, Monday to Friday, from 8:00 am to 5:30 pm.</p>

                <p>Kind regards,<br>
                <b>Joining team<br>
                TASC Group</b>
                </p>
            </body>
            </html>
            """
            sender_email = "nandhini2571999@gmail.com"  
            sender_password = "utfy ogzm qwks vfwt"  
            recipient_email = ["arti@tascoutsourcing.com","basha@tascoutsourcing.com","kokila.v@tascoutsourcing.com"]
            self.send_gmail_email(sender_email, sender_password, recipient_email, subject, html_body)

def main():
    obj = AppNotification()
    df = pd.read_csv(obj.path).fillna('')
    df['Activity End Date'] = pd.to_datetime(df['Activity End Date'])
    obj.get_residence_mail_notification(df)
    obj.get_mol_mail_notification(df)

if __name__ == '__main__':
    main()