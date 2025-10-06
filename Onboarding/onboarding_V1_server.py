import os
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import smtplib
from email import encoders
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage  # to load image while sending in email
import smtplib
from datetime import datetime
import time
import shutil
import logging 
from sqlalchemy import create_engine
from urllib.parse import quote_plus
import base64
from db_class import Database_Manager
from msal import ConfidentialClientApplication
from datetime import datetime, timedelta
import base64
import requests


server_credential = {
    "ip_server": "localhost",
    "username": "postgres",
    "password": "Happy@321#321"}
db = Database_Manager('TCS_Onboarding' ,server_credential)

username = quote_plus(server_credential['username']) 
password = quote_plus(server_credential['password']) 
host = server_credential['ip_server']
port = 5432
database = "TCS_Onboarding"

connection_string = f"postgresql+psycopg2://{username}:{password}@{host}:{port}/{database}"
engine = create_engine(connection_string)


log_filename = os.path.abspath(__file__)
logging.basicConfig(
    filename=log_filename.replace('.py', '.log'),
    filemode='a',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logging.info("TCS Onboarding Initiated. ")


class BrowserAutomation:
    def __init__(self, headless=False):
        options = webdriver.ChromeOptions()
        if headless:
            options.add_argument("--headless")
        service = Service(ChromeDriverManager().install())  
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.implicitly_wait(10)
        self.current_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        # Define your app credentials
        self.CLIENT_ID = "0b387074-4ba9-4304-bf36-36309318f0c9"
        self.TENANT_ID = "4395a262-82b5-4485-b45b-bfbf365f5bc3"
        self.CLIENT_SECRET = 'NJR8Q~QsMqrs9xUOOt4oYXKiwjxbUhSTxdZGrbm8'
        self.USERNAME = "adidas@tascoutsourcing.com"
        self.PASSWORD = "Q/466565691218oh"
        self.graphUserScopes = ["Mail.Read", "Mail.Send"]
        self.BASE_DIR = r"C:\Users\Baarecare\OneDrive - TASC LABOUR SERVICES L.L.C\PRO_Outsourcing"
        self.visa_app_dir = os.path.join(self.BASE_DIR, r"CLIENTS\UAE\Adidas\ADIDAS - Employee Documents\@EmpDirName@\Visa Application Docs")
        self.SAVE_DIR = os.path.join(self.BASE_DIR, r"CLIENTS\UAE\Adidas\ADIDAS - Employee Documents\@EmpDirName@\Govt Receipts & Invoice")
        self.is_browser_active = False
        
    def import_excel_to_db(self, excel_filename):

        df = pd.read_excel(excel_filename)
        df = df.fillna('')
        df_filtered = df[(df['Status'] != 'Under Process') & (df['Visa Process Start Date'] != '') & (df['Path'] != '')]
        df = df_filtered[['Status','Employee Name', 'Position', 'Contact Number', 'Email', 'Visa Process Start Date', 'Path','Typist EmailID']]
        df['Visa Process Start Date'] = pd.to_datetime(df['Visa Process Start Date'], errors='coerce').dt.strftime('%Y-%m-%d')
        df['EmployeeStatus'] = 'Active'
        df['ClientName'] = 'Adidas Emerging Markets L.L.C'
        df['TaskTitle'] = df['Employee Name'].apply(lambda x:x+'_New Visa Application')
        df['Project'] = 'New Visa Application (Mainland)'
        df['Assignee'] = 'Corporate Service Desk'
        df['Customer'] = 'Adidas Emerging Markets L.L.C'
        df['AddedDate'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        df['Subject'] = df.apply(lambda x: f"Job Offer & Contract typing - {x['Employee Name']} - ADIDAS", axis=1)
        df_renamed = df.rename(columns={
        'Status': 'tasc_stages',
        'Visa Process Start Date': 'start_date',
        'Employee Name': 'employee_name',
        'Position': 'job_position',
        'Contact Number': 'work_mobile',
        'Email': 'mail_id',
        'Path': 'file_path',
        'Typist EmailID': 'typist_email_id',
        'EmployeeStatus': 'employee_status',
        'ClientName': 'client_name',
        'TaskTitle': 'task_title',
        'Project':'project',
        'Assignee': 'assignee',
        'Customer': 'customer',
        'AddedDate': 'added_date',
        'Subject': 'subject',
        'FilePath': 'file_path'
    })

        df['FilePath'] = self.visa_app_dir
        # Forming filepath
        df['FilePath'] = df.apply(lambda row: row['FilePath'].replace('@EmpDirName@', row['Employee Name'].strip()), axis=1)
        if not df.empty:
            df_renamed = df_renamed.loc[:, ~df_renamed.columns.duplicated()]
            df_renamed.to_sql("tasc_tracker", engine, if_exists="append", index=False)
        else:
            print("There is no New records Found! Move to next")
        df['Status'] = 'Under Process'
        df.to_excel(excel_filename, index=False)
        self.df = df
        print("stage0 completed.")

    def go_to_url(self, url):
        self.driver.get(url)

    def login(self):
        login_element = self.driver.find_element(By.ID, "login")
        login_element.send_keys("projects@tascoutsourcing.com")
        time.sleep(1)
        password_element = self.driver.find_element(By.ID, "password")
        password_element.send_keys("Odoo@2025")
        time.sleep(1)
        submit_element = self.driver.find_element(By.XPATH,'//*[@id="wrapwrap"]/main/div[1]/form/div[3]/button')
        submit_element.click()
        print("Successfully login!")
        logging.info("Successfully login!")
        
    def close_driver(self):
        self.driver.close()

    def employee_entry(self, data):
        employee_element = self.driver.find_element(By.XPATH, '//*[@id="result_app_18"]/img')
        employee_element.click()
        time.sleep(20)
        try:
            self.driver.find_element(By.CSS_SELECTOR, "button.btn.btn-primary.o-kanban-button-new").click() 
        except:
            self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div/div[1]/div[1]/div[2]/button').click()
        time.sleep(2)
        self.driver.find_element(By.ID, 'name_0').send_keys(data['employee_name'])
        time.sleep(0.2)
        self.driver.find_element(By.ID, 'job_title_0').send_keys(data['job_position'])
        time.sleep(0.2)
        self.driver.find_element(By.ID, 'mobile_phone_0').send_keys(data['work_mobile'])
        time.sleep(0.2)
        self.driver.find_element(By.ID, 'work_email_0').send_keys(data['mail_id'])
        time.sleep(0.2)
        self.driver.find_element(By.ID, 'job_id_0').send_keys(data['job_position'])
        self.driver.find_element(By.ID, 'job_id_0_0_0').click()
        time.sleep(0.2)
        self.driver.find_element(By.ID, 'client_name_0').send_keys(data['client_name'])
        self.driver.find_element(By.ID, 'client_name_0_0_0').click()
        time.sleep(0.2)
        self.driver.find_element(By.ID, 'employee_status_0').send_keys(data['employee_status'])
        time.sleep(0.2)
        save_btns = ["/html/body/div[1]/div/div/div[1]/div/div[1]/div[3]/div/button[1]/i","/html/body/div[1]/div/div/div[1]/div/div[1]/div[3]/div/button[1]"]
        for save_btn in save_btns:
            try:
                self.driver.find_element(By.XPATH, save_btn).click()
                time.sleep(0.2)
                update_query = f"UPDATE tasc_tracker SET updated_date = '{self.current_date}', tasc_stages = 'Stage2 - New Employee Record Added' WHERE mail_id = '{data['mail_id']}'"
                try:
                    db.query_exec(update_query)  
                    logging.info(f"updated successfully - {update_query}")
                    data['tasc_stages'] = 'Stage2 - New Employee Record Added'
                    logging.info(f"status changing into : {data['tasc_stages']}")        
                    return data , True
                except: 
                    logging.info(f"Failed to update: {update_query}")
                break
            except Exception as e:
                update_query = f"UPDATE tasc_tracker SET updated_date = '{self.current_date}', tasc_stages = 'Failed' WHERE mail_id = '{data['mail_id']}'"
                try:
                    db.query_exec(update_query)  
                    logging.info(f"updated successfully - {update_query}")
                    data['TASCStatus'] = 'Failed'
                    logging.info(f"status changing into : {data['tasc_stages']}")        
                    return data, False
                    self.close_browser()
                except: 
                    logging.info(f"Failed to update: {update_query}")
        print("stage1 completed.")

    def assign_project(self, data):
        try:
            # Homepage - navbar click
            self.driver.find_element(By.XPATH, '/html/body/header/nav/a').click()
        except:
            pass
        time.sleep(2)
        self.driver.find_element(By.XPATH, '//*[@id="result_app_12"]/img').click()
        time.sleep(2)
        self.driver.find_element(By.XPATH, '/html/body/header/nav/div[1]/div[1]/button/span').click()
        time.sleep(2)
        self.driver.find_element(By.XPATH, '/html/body/header/nav/div[1]/div[1]/div/a[1]').click()
        time.sleep(5)
        # clicking NEW button
        new_button_paths = ['/html/body/div[1]/div/div[1]/div/div[1]/div[1]','/html/body/div[1]/div/div[1]/div/div[1]/div[1]/div[1]','/html/body/div[1]/div/div[1]/div/div[1]/div[1]/div[1]/button[1]','/html/body/div[1]/div/div[1]/div/div[1]/div[1]/div[2]']
        for new_button in new_button_paths:
            try:
                self.driver.find_element(By.XPATH, new_button).click()
                time.sleep(2)
                break
            except:
                pass 
        # task_title = 
        self.driver.find_element(By.ID, 'name_0').send_keys(data['employee_name'])
        # project = 
        self.driver.find_element(By.ID, 'project_id_1').send_keys(data['project'])
        self.driver.find_element(By.ID, 'project_id_1_0_0').click()
        # assignee = 
        self.driver.find_element(By.ID, 'user_ids_0').send_keys(data['assignee'])
        self.driver.find_element(By.ID, 'user_ids_0_0_0').click()
        # employee_name = 
        self.driver.find_element(By.ID, 'employee_id_0').send_keys(data['employee_name'])
        self.driver.find_element(By.ID, 'employee_id_0_0_0').click()
        # start_date = 
        self.driver.find_element(By.ID, 'task_planned_date_begin_0').send_keys(data['start_date'].strftime('%d/%m/%Y %H:00:00'))
        # save button click
        self.driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[1]/div/div[1]/div[3]/div/button[1]/i').click()
        try:
            db.query_exec(
                f"UPDATE tasc_tracker SET updated_date = '{self.current_date}', tasc_stages = 'Stage3 - Task Assigned' WHERE mail_id = '{data['mail_id']}'"
            )
            data['tasc_stages'] = 'Stage3 - Task Assigned'
            print('Stage 2 completed.')
            return data

        except: 
            print("failed to update the status, pls check. ")        
            logging.info("failed to update the status, pls check. ")   
            db.query_exec(
                f"UPDATE tasc_tracker SET updated_date = '{self.current_date}', tasc_stages = 'Failed' WHERE mail_id = '{data['mail_id']}'"
            )
            data['tasc_stages'] = 'Stage3 - Task Assigned'
            print('Stage 2 completed.')
            return data     
    
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

        # Attach HTML with CID reference
        msg_alternative.attach(MIMEText(body, 'html'))
        icon_paths = {
            "logo_cid": r"C:\Users\Baarecare\Automations\TCS_Onboarding\images\TASC_Logo.png",
            "phone_icon_cid": r"C:\Users\Baarecare\Automations\TCS_Onboarding\images\Phone.png",
            "mail_icon_cid": r"C:\Users\Baarecare\Automations\TCS_Onboarding\images\Mail.png",
            "location_icon_cid": r"C:\Users\Baarecare\Automations\TCS_Onboarding\images\Location.png",
            "website_icon_cid": r"C:\Users\Baarecare\Automations\TCS_Onboarding\images\Website.png",
            "bar_icon_cid": r"C:\Users\Baarecare\Automations\TCS_Onboarding\images\Bar.png",
            # Add more as needed
            }
        for cid, path in icon_paths.items():
            try:
                with open(path, "rb") as f:
                    image = MIMEImage(f.read())
                    image.add_header("Content-ID", f"<{cid}>")
                    image.add_header("Content-Disposition", "inline", filename=path.split("\\")[-1])
                    msg_root.attach(image)
            except Exception as e:
                print(f"Failed to attach {cid}: {e}")

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
                        logging.info(f"Attachment {attachment} added successfully.")
                except Exception as e:
                    print(e)
                    logging.error(f"Failed to attach {attachment}: {e}")

        # Send mail
        try:
            smtp_server = "smtp.office365.com"
            smtp_port = 587
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.sendmail(sender_email, all_recipients, msg_root.as_string())
                print(f"Email sent to {all_recipients}")
                logging.info(f"Email sent to {all_recipients}")
        except Exception as e:
            print(f"Failed to send email to {recipient_email}: {e}")
            logging.error(f"Failed to send email to {recipient_email}: {e}")

        # Update DB status
        db.query_exec(
            f"UPDATE tasc_tracker SET updated_date = '{self.current_date}', tasc_stages = 'Stage1 - typist Email Draft' WHERE mail_id = '{data['mail_id']}'"
        )
        data['tasc_stages'] = 'Stage1 - typist Email Draft'
        return data
        
    # Function to get all emails with pagination and date filtering
    def get_filtered_emails(self, headers, user_id, start_date):
        messages_url = f'https://graph.microsoft.com/v1.0/users/{user_id}/mailFolders/inbox/messages'
        filter_query = f"$filter=receivedDateTime ge {start_date.isoformat()}Z"
        messages_url = f'{messages_url}?{filter_query}'
        emails = []

        while messages_url:
            response = requests.get(messages_url, headers=headers)
            if response.status_code != 200:
                print(response.json())
                break
            response_data = response.json()
            emails.extend(response_data.get('value', []))
            messages_url = response_data.get('@odata.nextLink')
            print(messages_url)
        return emails
    
    def download_attachments(self, user_id, message_id , headers, data):
        is_attachment_downloaded = False
        attach_url = f"https://graph.microsoft.com/v1.0/users/{user_id}/messages/{message_id}/attachments"
        resp = requests.get(attach_url, headers=headers)
        if resp.status_code != 200:
            print(f"Failed to fetch attachments: {resp.json()}")
            return

        attachments = resp.json().get("value", [])
        for attachment in attachments:
            try:
                if attachment["@odata.type"] == "#microsoft.graph.fileAttachment" and attachment['contentType'] == 'application/pdf':
                    file_name = attachment["name"]
                    content_bytes = attachment["contentBytes"]
                    file_data = base64.b64decode(content_bytes)
                    file_path = self.SAVE_DIR.replace('@EmpDirName@',data['employee_name'])
                    os.makedirs(file_path, exist_ok=True)
                    full_file_path = os.path.join(file_path, file_name)
                    with open(full_file_path, "wb") as f:
                        f.write(file_data)
                    print(f"Downloaded: {file_path}")
                    logging.info(f"Downloaded: {file_path}")
                    is_attachment_downloaded = True
            except Exception as e:
                print(e)
                is_attachment_downloaded = False 
        
        return is_attachment_downloaded 

    def move_files_to_tracker(self,attachment_dir: str, tracker_folder_name: str = 'Typed Documents'):
        typed_dir = attachment_dir.replace('Govt Receipts & Invoice',tracker_folder_name)

        # Ensure the Tracker folder exists
        os.makedirs(typed_dir, exist_ok=True)

        moved_files_count = 0
        skipped_files_count = 0

        try:
            # List all files in the attachment directory
            for filename in os.listdir(attachment_dir):
                file_path = os.path.join(attachment_dir, filename)

                # Skip directories, process only files
                if os.path.isfile(file_path):
                    # Split filename into name and extension
                    name_without_ext, _ = os.path.splitext(filename)

                    # Check if the name ends with 'MB' or 'ST'
                    if name_without_ext.endswith('MB') or name_without_ext.endswith('ST'):
                        destination_path = os.path.join(typed_dir, filename)
                        try:
                            shutil.move(file_path, destination_path)
                            print(f"Moved: '{filename}' to '{typed_dir}'")
                            moved_files_count += 1
                        except Exception as e:
                            print(f"Error moving '{filename}': {e}")
                    else:
                        skipped_files_count += 1

        except FileNotFoundError:
            print(f"Error: Source directory '{attachment_dir}' not found.")
            return
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            return
        
    def check_typist_email(self, sender_email, password, data):
        # Create a confidential client application
        app = ConfidentialClientApplication(
            self.CLIENT_ID,
            authority=f'https://login.microsoftonline.com/{self.TENANT_ID}',
            client_credential=self.CLIENT_SECRET,
        )

        # Acquire a token for the client
        token_response = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        access_token = token_response.get('access_token')

        # Define the headers with the acquired access token
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        # Define the start date for filtering emails
        start_date = datetime.now() - timedelta(days=10)  # Example: 30 days ago

        # Get all emails from the inbox for the specified user and date filter
        emails = self.get_filtered_emails(headers, self.USERNAME, start_date)

        if emails:
            logging.info("Got response to search unseen emails")
            logging.info("*** Process Starts! ***")
        

        for email in emails:
            fields = {}
            try:
                if email['subject'].lower() == 're: '+data['subject'].lower():
                    logging.info(f"RE: {data['subject']} found and started searching for attachments.")
                    print("found the mail.")
                    fields['msgdate'] = email['receivedDateTime'].replace('T', ' ').replace('Z', '')
                    fields['msgto'] = email['toRecipients'][0]['emailAddress']['address']
                    fields['msgfrom'] = email['from']['emailAddress']['address']
                    fields['msgsubject'] = email['subject']
                    fields['msgfromname'] = email['from']['emailAddress']['name']
                    fields['Source'] = '1'
                    fields['htmltag'] = email['body']['content']
                    fields['body'] = re.sub(r'\s+', ' ', email['bodyPreview'].replace('\n', '').replace('\r', '').strip())
                    fields['msgid'] = email['internetMessageId']
                    fields['msguid'] = email['conversationId']    
                    # print(fields)
                    logging.info(f"Is Attachment available: {email.get('hasAttachments')}")
                    # Download attachments if present
                    if email.get('hasAttachments'):
                        message_id = email['id']
                        is_attachment_downloaded = self.download_attachments(self.USERNAME, message_id , headers, data)
                        if is_attachment_downloaded:
                            # Moving the _ST and _MD files to Tracker 
                            file_path = self.SAVE_DIR.replace('@EmpDirName@',data['employee_name'])
                            self.move_files_to_tracker(file_path, 'Typed Documents')
                            print("Email downloaded successfully.")
                            db.query_exec(
                                f"UPDATE tasc_tracker SET updated_date = '{self.current_date}', tasc_stages = 'Stage4 - Reply from typist' WHERE mail_id = '{data['mail_id']}'"
                            )
                            data['tasc_stages'] = 'Stage4 - Reply from typist'
                            return data
                        else:
                            logging.info("failed to download the attachment.")
                            db.query_exec(f"update tasc_tracker set updated_date = '{self.current_date}', tasc_stages = 'Failed to download the attachment' WHERE mail_id = '{data['mail_id']}'")
                            return data
            except Exception as e:
                print("email : ", email)
                logging.info(f"Extracting Issue With This Mail Content: {email}, Error: {e}")

        logging.info("*** Process Completed ***")
        print("*** Process Completed ***")

    def send_employee_email(self, sender_email, sender_password, employee_email, body, data):       
        try:
            attachment_dir = self.SAVE_DIR.replace('@EmpDirName@',data['employee_name']).replacanye('Govt Receipts & Invoice','Typed Documents')
            msg_root = MIMEMultipart('related')
            msg_root['From'] = sender_email
            msg_root['To'] = employee_email
            # msg_root['To'] = 'veena.a@tascoutsourcing.com'
            msg_root['Cc'] = ''
            msg_root['Bcc'] = ''
            all_recipients = [msg_root['To']] + msg_root['Cc'].split(',') + msg_root['Bcc'].split(',')
            msg_root['Subject'] = data['subject']
            msg_root.attach(MIMEText(body,'html'))
            attachment_paths = [os.path.join(attachment_dir, i) for i in os.listdir(attachment_dir) if i.split('.')[0].endswith('MB') or i.split('.')[0].endswith('ST')]

            # HTML body part
            msg_alternative = MIMEMultipart('alternative')
            msg_root.attach(msg_alternative)

            # Attach HTML with CID reference
            msg_alternative.attach(MIMEText(body, 'html'))
            icon_paths = {
                "logo_cid": r"C:\Users\Baarecare\Automations\TCS_Onboarding\images\TASC_Logo.png",
                "phone_icon_cid": r"C:\Users\Baarecare\Automations\TCS_Onboarding\images\Phone.png",
                "mail_icon_cid": r"C:\Users\Baarecare\Automations\TCS_Onboarding\images\Mail.png",
                "location_icon_cid": r"C:\Users\Baarecare\Automations\TCS_Onboarding\images\Location.png",
                "website_icon_cid": r"C:\Users\Baarecare\Automations\TCS_Onboarding\images\Website.png",
                "bar_icon_cid": r"C:\Users\Baarecare\Automations\TCS_Onboarding\images\Bar.png",
                # Add more as needed
            }
            

            for cid, path in icon_paths.items():
                try:
                    with open(path, "rb") as f:
                        image = MIMEImage(f.read())
                        image.add_header("Content-ID", f"<{cid}>")
                        image.add_header("Content-Disposition", "inline", filename=path.split("\\")[-1])
                        msg_root.attach(image)
                except Exception as e:
                    print(f"Failed to attach {cid}: {e}")
                
            if attachment_dir:
                for attachment in attachment_paths:
                    try:
                        file_name = attachment.split('\\')[-1]
                        if file_name.split('.',1)[0].endswith('_MB') or file_name.split('.',1)[0].endswith('_ST'):
                                
                            with open(attachment, 'rb') as file:
                                part = MIMEBase('application', 'octet-stream')
                                part.set_payload(file.read())
                                encoders.encode_base64(part)  # Encode file in base64
                                part.add_header('Content-Disposition', f'attachment; filename={file_name}')
                                msg_root.attach(part)
                                logging.info(f"Attachment {attachment} added successfully.")
                    except Exception as e:
                        print(e)
                        logging.error(f"Failed to attach {attachment}: {e}")
            try:
                smtp_server = "smtp.office365.com"
                smtp_port = 587
                with smtplib.SMTP(smtp_server, smtp_port) as server:
                    server.starttls()
                    server.login(sender_email, sender_password)
                    logging.info(f"Mail sent to : {all_recipients}")
                    server.sendmail(sender_email, all_recipients, msg_root.as_string())
                    print(f"Email sent to {all_recipients}")
                    logging.info(f"Email sent to {all_recipients}")
                
            except Exception as e:
                print(f"Failed to send email to {recipient_email}: {e}")
                logging.info(f"Failed to send email to {recipient_email}: {e}")
            print("Email sent successfully.")
            db.query_exec(
                f"UPDATE tasc_tracker SET updated_date = '{self.current_date}', tasc_stages = 'Stage5 - Employee Email Draft' WHERE mail_id = '{data['mail_id']}'"
            )
            data['tasc_stages'] = 'Stage5 - Employee Email Draft'
            return data
        
        except Exception as e:
            logging.info(f"failed to send email: {e}")
            
    def get_page_title(self):
        return self.driver.title

    def close_browser(self):
        self.driver.quit()



if __name__ == "__main__":
    bot = BrowserAutomation()  
    try:
        excel_filename = r'C:\Users\BotServer_2.WIN-IJ6UPG7E58S\OneDrive - TASC LABOUR SERVICES L.L.C\PRO_Outsourcing - TCS AUTOMATIONS\Tracker - TASC Applications2025_test1.xlsx'
        # excel_filename = r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\Onboarding\Tracker - TASC Applications2025_test1.xlsx'#Duplicate
        df = bot.import_excel_to_db(excel_filename)
        try:
            details = db.fetchArray_withKey("select * from public.TASC_Tracker where tasc_stages <> 'Stage5 - Employee Email Draft' order by id")

        except Exception as e: 
            logging.info(f"Some Error as occured while fetching the data : {e}")
            details = dict()
        for data in details:
            sender_email = "adidas@tascoutsourcing.com"  
            sender_password = "Q/466565691218oh"
            employee_email = data['mail_id']
            recipient_email = data['typist_email_id']   
            body = "Please find the updated document attached."
            body = r"""
            <html><body>
            <p>Hi Team, <br><br>
            Please proceed with drafting the job offer and contract. Attached are the required documents. <br><br>
            <b>Imp Note: Kindly save the MB and ST file name as (@EmployeeName@_MB and @EmployeeName@_ST) <br><br>
            <h3>Company Name: ADIDAS EMERGING MARKETS L.L.C<br>
            Company Code: 216152</h3>
            <br>
            <table border=0>
            <tr><td style="padding-right: 100px;"><b>Employee Name</b><td style="padding-right: 10px;">:<td>@EmployeeName@</tr>
            <tr><td style="padding-right: 100px;"><b>Profession</b><td style="padding-right: 10px;">:<td>@VisaDesignation@</tr>
            <tr><td style="padding-right: 100px;"><b>Phone number</b><td style="padding-right: 10px;">:<td>@data_to_be_picked_from_Online_tracker_ph@</tr>
            <tr><td style="padding-right: 100px;"><b>Email ID</b><td style="padding-right: 10px;">:<td>@data_to_be_picked_from_Online_tracker_email@</tr>
            <tr><td> </td><tr>
            <tr><td> </td><tr>
            <tr><td style="padding-right: 20px;"><b>Salary Details</b><td style="padding-right: 10px;">:<td>As per attached Offer letter</tr>
            <tr><td style="padding-right: 20px;"><b>Marital status & Religion</b><td style="padding-right: 10px;">:<td>As mentioned in the attached Personal information sheet</tr>
            <tr><td style="padding-right: 20px;"><b>Probation</b><td style="padding-right: 10px;">:<td>6 months</tr>
            <tr><td style="padding-right: 20px;"><b>Notice Period</b><td style="padding-right: 10px;">:<td>30 days</tr>
            <tr><td style="padding-right: 20px;"><b>Weekly Day off</b><td style="padding-right: 10px;">:<td>1 (random days except weekend)
            </table>
            <br><br>
            <b>Best Regards,</b>
            <br><br>
            <table width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
                <!-- Left: Company Logo (30%) -->
                <td align="left" style="width: 30%;">
                <img src="cid:logo_cid" alt="Company Logo" style="height: 60px;">
                </td>

                <!-- Center: Signature Details (30%) -->
                <td align="left" style="width: 30%; font-size: 14px; line-height: 24px; color: #ffffff; background-color: #1e1e1e; padding: 10px;">
                
                <div style="font-weight: bold;align=center;">Veena Arul</div>
                <div><img src="cid:bar_icon_cid" style="height: 20px;"></div>
                <div><img src="cid:phone_icon_cid" style="height: 20px;">
                    &nbsp;+971 52 9299509
                </div>
                <div><img src="cid:mail_icon_cid" style="height: 20px;">
                    <a href="mailto:adidas@tascoutsourcing.com" style="color: #1e90ff; text-decoration: none;">&nbsp;adidas@tascoutsourcing.com</a>
                </div>
                <div><img src="cid:location_icon_cid" style="height: 20px;">&nbsp;Business Central Towers, Tower A  - 1506 – Dubai Internet City
                </div>
                <div><img src="cid:website_icon_cid" style="height: 20px;">
                    <a href="www.tasccorporateservices.com" style="color: #1e90ff; text-decoration: none;">&nbsp;www.tasccorporateservices.com</a>
                </div>
                </td>
                <!-- Right: Blank Space (40%) -->
                <td style="width: 40%;">&nbsp;</td> 
                </tr>
                </table>
                <br><br>
                <hr style="border: 3; height: 1px; background-color: #000; margin: 10px 0;">
                </body></html>
            """
            body = body.replace('@EmployeeName@',data['employee_name']).replace('@VisaDesignation@',data['job_position']).replace('@data_to_be_picked_from_Online_tracker_ph@',data['work_mobile']).replace('@data_to_be_picked_from_Online_tracker_email@',data['mail_id'])
            attachment_dir = [os.path.join(data['file_path'], file) for file in os.listdir(data['file_path'])]
            try:
                if data['tasc_stages'] == 'New':
                    logging.info("Processing Stage0 started.")
                    data = bot.send_typist_email(sender_email, sender_password, recipient_email, body, data, attachment_dir= attachment_dir)
                    logging.info("Processing Stage0 Ended.")
            except Exception as e:
                print(e)
            try:
                if data['tasc_stages'] == 'Stage1 - typist Email Draft':
                    logging.info("Processing Stage1 started")
                    bot.go_to_url("https://tascodoo-tascoutsourcing.odoo.com/web/login")
                    time.sleep(2)  
                    print("Page Title:", bot.get_page_title())
                    bot.login()
                    data, bot.is_browser_active =bot.employee_entry(data)
                    logging.info("Processing Stage1 started")
                    bot.df.loc[bot.df['Email'] == data['mail_id'], 'Status'] = 'Under Process'
            except Exception as e:
                print(e)
            try:
                if data['tasc_stages'] == 'Stage2 - New Employee Record Added':
                    logging.info("Processing Stage2 started")
                    if not bot.is_browser_active:
                        bot.go_to_url("https://tascodoo-tascoutsourcing.odoo.com/web/login")
                        time.sleep(2)  
                        print("Page Title:", bot.get_page_title())
                        bot.login()
                    data = bot.assign_project(data)
                    try:
                        bot.close_driver()
                    except Exception as e:
                        logging.info("unable to close the driver.")
                        print("unable to close the driver.")
                    logging.info("Processing Stage2 ended.")
            except Exception as e:
                print(e)
            try:
                if data['tasc_stages'] == 'Stage3 - Task Assigned':
                    logging.info("Processing Stage3 started")
                    emp_attachment_dir = bot.check_typist_email(sender_email, sender_password, data)
                    logging.info("Processing Stage3 ended.")
                
            except Exception as e:
                print(e) 
            try:
                if data['tasc_stages'] == 'Stage4 - Reply from typist':
                    logging.info("Processing Stage4 started")
                    body = rf"""<html><body>
                    <p><br><br>
                    Dear {data['employee_name']},
                    <br><br> 
                    I hope this message finds you well.
                    <br><br> 
                    Please find attached your <b>Job Offer and Employment Contract</b> for your review and signature. To proceed with your application, kindly complete the following steps:<br><br>
                    <ul>
                    <li>Review, <b>sign, and return the attached documents</b> at your earliest convenience.</li>
                    <li><b>Sign the MB and ST contracts via the email link</b> sent by <b>MOHRE on your personal email</b>. Once completed, please confirm by replying to this email.</li>
                    <li>Confirm whether you are currently inside or outside the UAE.</li>
                    </ul>
                    <br>
                    <b>Next Steps:</b>  
                    Upon receipt of the above, we will initiate the required approvals with MOHRE (Ministry of Human Resources and Emiratisation) and proceed with visa processing through Immigration. We will keep you updated on all progress and notify you as soon as approvals are obtained.
                    <br><br> 
                    We look forward to your prompt response and completion of the steps mentioned above.
                    <br><br>
                    <b>Best Regards,</b>
                <br><br>
                <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                    <!-- Left: Company Logo (30%) -->
                    <td align="left" style="width: 30%;">
                    <img src="cid:logo_cid" alt="Company Logo" style="height: 60px;">
                    </td>

                    <!-- Center: Signature Details (30%) -->
                    <td align="left" style="width: 30%; font-size: 14px; line-height: 24px; color: #ffffff; background-color: #1e1e1e; padding: 10px;">
                    <div style="font-weight: bold;align=center;">Veena Arul</div>
                    <div><img src="cid:bar_icon_cid" style="height: 20px;"></div>
                    <div><img src="cid:phone_icon_cid" style="height: 20px;">
                        &nbsp;+971 52 9299509
                    </div>
                    <div><img src="cid:mail_icon_cid" style="height: 20px;">
                        <a href="mailto:adidas@tascoutsourcing.com" style="color: #1e90ff; text-decoration: none;">&nbsp;adidas@tascoutsourcing.com</a>
                    </div>
                    <div><img src="cid:location_icon_cid" style="height: 20px;">
                        &nbsp;Business Central Towers, Tower A  - 1506 – Dubai Internet City
                    </div>
                    <div><img src="cid:website_icon_cid" style="height: 20px;">
                        <a href="www.tasccorporateservices.com" style="color: #1e90ff; text-decoration: none;">&nbsp;www.tasccorporateservices.com</a>
                    </div>
                    </td>
                    <!-- Right: Blank Space (40%) -->
                    <td style="width: 40%;">&nbsp;</td> 
                </tr>
                </table>
                <br><br>
                <hr style="border: 3; height: 1px; background-color: #000; margin: 10px 0;">
                </body></html>
                """
                    data = bot.send_employee_email(sender_email, sender_password, employee_email, body, data)
                    logging.info("Processing Stage4 ended.")
            except Exception as e:
                print(f"Exception as occured: {e}")  
            # try:  
            #     if data['tasc_stages'] != 'New':
            #         update_query = f"UPDATE tasc_tracker SET tasc_stages = 'Under Process' WHERE mail_id = '{data['mail_id']}'"
            #         try:
            #             db.query_exec(update_query)
            #         except Exception as e:
            #             logging.info(f"Failed to update the status : {e}")
            #         data['TASCStatus'] = 'Under Process'
            #         bot.df.loc[bot.df['Email'] == data['mail_id'], 'Status'] = 'Under Process'
            #         # bot.df.to_excel(r"C:\Users\Baarecare\OneDrive - TASC LABOUR SERVICES L.L.C\PRO_Outsourcing\TCS AUTOMATIONS\Tracker - TASC Applications 2025_test.xlsx", index=False)
            #     print("success")
            # except Exception as e:
            #     print(e)
    except Exception as e:
        print("The exception as occured: ", e)

    finally:
        bot.close_browser()

"""
Attachment like ST, MB
""" 
