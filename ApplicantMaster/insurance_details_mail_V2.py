import os
import json
import requests
import pandas as pd
import smtplib
from email.message import EmailMessage
from bs4 import BeautifulSoup
from datetime import datetime
from dateutil.parser import parse
import base64
import shutil
import pytz
local_tz = pytz.timezone("Asia/Dubai") 

class ApplicantMaster:
    def __init__(self):
        self.file_path = r"Z:\Insurance Files\Insurance Card Copies\Applicant_Master\Applicant_report.csv"
        self.file_path = r"C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\ApplicantMaster\Applicant_report_05064210062025.csv"
        self.card_dir = r"Z:\Insurance Files\Insurance Card Copies\Bot"
        self.card_dir = r"C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\ApplicantMaster"
        self.card_details = r"C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\ApplicantMaster\CardDetails.xlsx"
    
    def detect_card_type(self, text):
        if 'ADNIC' in text:
            return 'type1'
        elif 'TASC OUTSOURCING' in text:
            return 'type2'
        elif 'HEALTHCARE' in text:
            return 'type3'
        elif 'Select Silver' in text:
            return 'type4'
        elif 'MetLife' in text:
            return 'type5'
        else:
            return None

    def date_format(self,date_str):
        for format in ['%Y-%m-%d', '%d-%b-%Y','%d/%m/%Y','%d%m%Y','%d/%m%Y']:
            try:
                date_str = datetime.strptime(date_str, format).strftime('%Y-%m-%d')
                return date_str
            except:
                pass
        return ''

    def insurance_email_draft(self, sender_email, sender_password, recipient_email, subject, body, attachment_dir, data):
        self.email_status = False
        msg = EmailMessage()
        msg["From"] = sender_email
        msg["To"] = recipient_email
        msg["Subject"] = subject
        msg.set_content(body)
    
        try:
            # Find employee image attachments
            attachment_paths = [
                os.path.join(attachment_dir, i)
                for i in os.listdir(attachment_dir)
                if i.startswith(str(data['EmpCode'])) and (i.endswith('jpg') or i.endswith('png'))
            ]

            if str(data['Work Company']) == 'TASC Labour Services LLC':
                root_dir = os.path.join(
                    r'C:\Users\Developer\OneDrive - TASC LABOUR SERVICES L.L.C',
                    "Request Tasc's files - All Policies",
                    '1000-TASC Labour Services LLC'
                )
            elif str(data['Work Company']) == 'Top Talent Employment services LLC':
                root_dir = os.path.join(
                    r'C:\Users\Developer\OneDrive - TASC LABOUR SERVICES L.L.C',
                    "Request Tasc's files - All Policies",
                    '2000-Top Talent Employment Services LLC'
                )
            else:
                raise ValueError(f"Unknown Work Company: {data['Work Company']}")

            # Verify the directory exists
            if not os.path.exists(root_dir):
                raise FileNotFoundError(f"Directory not found: {root_dir}")

            # List directories matching MedicalInsurance
            directory = [i for i in os.listdir(root_dir) if data['MedicalInsurance'] == i]
            if directory:
                insurance_dir = directory[0]
                main_dir = os.path.join(root_dir, insurance_dir)
                insurance_docs = [os.path.join(main_dir, i) for i in os.listdir(main_dir)]
    
                final_docs = attachment_paths + insurance_docs
    
                for attachment_path in final_docs:
                    with open(attachment_path, "rb") as f:
                        file_data = f.read()
                        file_name = os.path.basename(attachment_path)
                        msg.add_attachment(
                            file_data,
                            maintype="application",
                            subtype="octet-stream",
                            filename=file_name,
                        )
    
                # Send email using Outlook SMTP
                with smtplib.SMTP("smtp.office365.com", 587) as server:
                    server.starttls()  # Use TLS (not SSL)
                    server.login(sender_email, sender_password)
                    server.send_message(msg)
                    self.email_status = True
                print("Email sent successfully via Outlook.")
                data['TASCStatus'] = 'Success'
            else:
                print("Medical Insurance Docs Not Found")
                data['TASCStatus'] = 'Folder Name Not Matched'
    
        except Exception as e:
            print(f"Failed to send email: {e}")
            data['UploadStatus'] = 'Failed'
    
    def convert_bas64(self, img):   
        with open(img, "rb") as f:
            encoded = base64.b64encode(f.read()).decode("utf-8")
            return encoded

    def convert_dt_milliseconds(self, date_str):
        localized_date = local_tz.localize(parse(date_str).replace(hour=20, minute=0, second=0, microsecond=0))
        unix_timestamp = int(localized_date.timestamp() * 1000)
        return unix_timestamp
    
    def upload_api_data(self, final_data):
        try:
            self.img_file_path = [os.path.join(self.card_dir, i) for i in os.listdir(self.card_dir) if i.startswith(str(final_data['EmpCode']))][0]
            with open(self.img_file_path, "rb") as image_file:
                encoded_string = base64.b64encode(image_file.read()).decode("utf-8")

            url = "https://api22.sapsf.com/odata/v2/upsert?"
            username = "SFAPI@tasclabour"
            password = "Welcome111"

            payload = {
                "__metadata": {"uri": "Attachment"},
                "fileName": os.path.basename(self.img_file_path),
                "module": "GENERIC_OBJECT",
                "userId": str(final_data['EmpCode']),
                "viewable": True,
                "fileContent": encoded_string
            }

            response = requests.post(
                url,
                auth=(username, password),
                headers={"Content-Type": "application/json"},
                data=json.dumps(payload)
            )
            if response.status_code == 200 and 'ERROR' not in response.text:
                soup = BeautifulSoup(response.content, 'xml')
                key_tag = soup.find('key')
                today = datetime.today()
                if key_tag:
                    attachment_id = key_tag.text.split("=")[-1]
                    doc_payload  = {
                        "__metadata": {
                            "uri": "cust_DocIDInfowExpiry",
                            "type": "SFOData.cust_DocIDInfowExpiry"
                        },
                        "effectiveStartDate": f"/Date({self.convert_dt_milliseconds(today)})/",
                        "externalCode": str(final_data['EmpCode']),
                        "cust_DocDetails": {
                            "cust_DocumentNumber":final_data['CardNo'],
                            "cust_Document_Type": "Ec_ARE_26",
                            "cust_IssueDate": f"/Date({self.convert_dt_milliseconds(final_data['EffectiveDate'])})/",
                            "cust_ExpiryDate": f"/Date({self.convert_dt_milliseconds(final_data['ExpiryDate'])})/",
                            "cust_attachmentNav": {
                                "__metadata": {
                                    "uri": f"Attachment(attachmentId={attachment_id})"
                                }
                            }
                        }
                    }
                    doc_response = requests.post(url,auth=(username, password),headers={"Content-Type": "application/json"},data=json.dumps(doc_payload))
                    if doc_response.status_code == 200 and 'ERROR' not in doc_response.text:
                        final_data['UploadStatus'] = 'Success'
                    else:
                        final_data['UploadStatus'] = '2API-Failed'
            else:
                final_data['UploadStatus'] = '1API-Failed'

        except Exception as e:
            print(f"Error Occured: {e}")
            final_data['UploadStatus']  = '1API-Failed'

    def move_img_file_dir(self, data):
        try:
            if data['TASCStatus'] == 'Success':
                shutil.move(self.img_file_path, r'Z:\Insurance Files\Insurance Card Copies\Bot\Done')
            else:
                shutil.move(self.img_file_path, r'Z:\Insurance Files\Insurance Card Copies\Bot\Error')
        except Exception as e:
            print(f"Error Occured: {e}")

    def fetch_card_information(self, code):
        df = pd.read_excel(self.card_details)
        df_filter = df[df['EmpCode'] == code]
        if not df_filter.empty:
            return {i: j[0] for i, j in df_filter.to_dict().items()}
        else:
            return {}
        
    def collect_emp_info(self):
        df = pd.read_csv(self.file_path)
        employee_code = [int(str(i).split('.')[0]) for i in os.listdir(self.card_dir) if str(i).split('.')[0].isdigit()]
        print(employee_code)
        content_list = []
        for code in employee_code:
            df_filter = df[df['ApplicantId'] == code]
            final_data = {'MedicalInsurance':'','CandidateEmailID':'', 'CardNo':'', 'ExpiryDate':'', 'EffectiveDate':'', 'UploadStatus':'', 'ExtractionStatus':'','EmailStatus':'N', 'TASCStatus': 'Completed'}
            final_data['EmpCode'] = code
            
            if not df_filter.empty:
                df_filter = df_filter.rename(columns={'ApplicantId': 'EmpCode', 'Medical Insurance':'MedicalInsurance', 'Candidate Email ID':'CandidateEmailID'})
                row = df_filter[['EmpCode', 'MedicalInsurance', 'Work Company', 'CandidateEmailID']].to_dict(orient='records')[0]
                if '/' in row['MedicalInsurance']:
                    final_data['MedicalInsurance'] = row['MedicalInsurance'].split('/')[2]
                    final_data['CandidateEmailID'] = row['CandidateEmailID']
                    final_data['Work Company'] = row['Work Company']
                    sender_email = "cards@tascoutsourcing.com"  
                    sender_password = "Loc52442"  
                    # recipient_email = "kokila.v@tascoutsourcing.com" 
                    recipient_email = final_data['CandidateEmailID'] 
                    # print(recipient_email) 
                    # recipient_email = final_data['CandidateEmailID']
                    sender_email = "nandhini2571999@gmail.com"  
                    sender_password = "utfy ogzm qwks vfwt"  
                    recipient_email = "nandylokesh123@gmail.com" #final_data['CandidateEmailID']  
                    subject = "Medical Insurance Card-" + str(row['EmpCode'])
                    print(final_data['MedicalInsurance'])
                    if 'ADNIC' in final_data['MedicalInsurance']:
                        # "ADNIC"
                        message_body = '''Dear Employee,

Kindly refer to attached files for your Table of Benefits and Network List of the Medical Insurance Policy that you are currently enrolled in. Note that ADNIC is cardless, hence you may use your Emirates ID to avail the medical treatment.

Please note that the network list is subject to change at times, hence we highly advise to view the updated Network List via the App to ensure that you visit an In-Network facility. Outside network treatments will only be reimbursed by the insurance company depending on the claims settlement of your policy (Refer to the TOB)

Also note that you are eligible for TruDoc Teleconsultation benefit. Please refer to the attached flyer for a short overview on the features and benefits.

Regards
Cards Team
'''
                    elif 'QIC' in final_data['MedicalInsurance']:
                        # "QIC"
                        message_body = '''Dear Employee,

Kindly refer to attached files for your Table of Benefits and Network List of the Medical Insurance Policy that you are currently enrolled in. Note that QIC is cardless, hence you may use your Emirates ID to avail the medical treatment.

Please note that the network list is subject to change at times, hence we highly advise to view the updated Network List via the App to ensure that you visit an In-Network facility. Outside network treatments will only be reimbursed by the insurance company depending on the claims settlement of your policy (Refer to the TOB)

Also note that you are eligible for TruDoc Teleconsultation benefit. Please refer to the attached flyer for a short overview on the features and benefits.            

Regards
Cards Team
'''
                    elif 'Sukoon' in final_data['MedicalInsurance']:
                        # "Sukoon"
                        message_body = '''Dear Employee,

Kindly find the insurance details for the policy you are enrolled in.

We kindly suggest you register on the Sukoon App for easy access to your medical insurance details and generate your E-Card copy since there will be no physical card available. 

Please note that the network list is subject to change at times, hence we highly advise to view the updated Network List via the App to ensure that you visit an In-Network facility. Outside network treatments will only be reimbursed by the insurance company depending on the claims settlement of your policy (Refer to the TOB)

Also note that you are eligible for TruDoc Teleconsultation benefit. Please refer to the attached flyer for a short overview on the features and benefits.

Please feel free to reach out to us if you have any queries. 

Regards 
Cards Team
'''
    
                    elif 'Daman' in final_data['MedicalInsurance']:
                        # "DAMAN"
                        message_body = '''Dear Employee,

Please see the below details related to your Daman Medical Insurance Policy:

1) Please find the attached relevant documents for your medical Insurance Policy
    •	Table Of Benefits
    •	Network List
    •	General Exclusion List

2) Attached also is the Daman Mobile App user guide for accessing the below services:
    • Access digital card details, insurance benefits, network list and policy details for employees and your insured dependents.

3) Please refer to the link below to access the list of international providers who will accept direct billing in case you are outside UAE.
    •	Please find link and choose outside UAE facility.  Link : Search Facility

3) Toll-free number for insurance related queries.
    •	The toll-free number is 600 5 insurance_email_draft

4) Teleconsultation process:
    •	Teleconsultation is available through Prime medical centers. Please find attached a list of providers for teleconsultation. You can either call the provider or book from their mobile application and inquire about the issue and the provider will call them back for the teleconsultation booking.

Thanks and Regards,
Cards Team
'''
                    elif 'MetLife' in final_data['MedicalInsurance']:
                        # "MetLife"
                        message_body = '''Dear Employee,

Kindly refer to attached files for the e-medical card, Table of Benefits (TOB) and Network List of the policy that you are enrolled in. Note that Metlife is cardless, hence you may use your Emirates ID to avail the medical treatment.

Please note that the network list is subject to change at times, hence we highly advise to view the updated Network List via the App/portal to ensure that you visit an In-Network facility. Outside network treatments will only be reimbursed by the insurance company depending on the claims settlement of your policy (Refer to the TOB)

Regards
Cards Team

'''
                    print(message_body)
                    if message_body:
                        attachment_dir = self.card_dir
                        self.insurance_email_draft(sender_email, sender_password, recipient_email, subject, message_body, attachment_dir, final_data)
                        self.email_status = True
                        final_data['EmailStatus'] = 'Y' if self.email_status == True else 'N'
                else:
                    print(f"There is no Medical Insurance Details")
                    final_data['TASCStatus'] = 'InsuranceDetail Not in the ApplicantMaster'
            else:
                print(f"ApplicanID: {code} Not Found in Excel")
                final_data['TASCStatus'] = 'ApplicantID Not in the ApplicanMaster'
            # record2 = self.extract_card_information("10128711")
            record2 = self.fetch_card_information(code)
            if record2:
                final_data = {**final_data, **record2}
                final_data['ExtractionStatus'] = 'Success'
                if final_data['ExtractionStatus'] == 'Success':
                    self.upload_api_data(final_data)
                    if final_data['UploadStatus'] == 'Success':
                        self.move_img_file_dir(final_data)
            else:
                final_data['ExtractionStatus'] = 'Failed'
            content_list.append(final_data) 
        self.generate_logfile(content_list)
        print("Log Generated Successfully")  
    
    def generate_logfile(self, records):
        df = pd.DataFrame(records)
        df = df[['image_resolution','EmpCode', 'CandidateEmailID', 'CardNo', 'ExpiryDate','EffectiveDate','MedicalInsurance', 'EmailStatus', 'ExtractionStatus' , 'UploadStatus' , 'TASCStatus']]
        timestamp = (datetime.now()).strftime('%Y%m%d%H%M%S')

        # df.to_excel(f"Z:\\Insurance Files\\Insurance Card Copies\\Log\\Log_{timestamp}.xlsx", index=False)
        df.to_excel(f"C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\ApplicantMaster\\Log_{timestamp}.xlsx", index=False)
if __name__ == "__main__":
    app = ApplicantMaster()
    app.collect_emp_info()