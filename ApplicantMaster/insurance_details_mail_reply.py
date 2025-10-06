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
import easyocr
# from paddleocr import PaddleOCR
from PIL import Image
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
import pytz
local_tz = pytz.timezone("Asia/Dubai") 

class ApplicantMaster:
    def __init__(self):
        self.file_path = r"C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\ApplicantMaster\Applicant_report_05064210062025.csv"
        self.card_dir = r"C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\ApplicantMaster"

    def get_image_resolution(self):
        # importing the module
        import PIL
        from PIL import Image

        # loading the image
        img = PIL.Image.open(self.image_path)

        # fetching the dimensions
        wid, hgt = img.size

        # displaying the dimensions
        return str(wid) + "x" + str(hgt)

    def tesseract_text(self):
        import cv2
        import easyocr
        from PIL import Image
        import pytesseract
        import numpy as np
        img = cv2.imread(self.image_path)
        upscaled = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
        gray = cv2.cvtColor(upscaled, cv2.COLOR_BGR2GRAY)
        clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(gray)
        denoised = cv2.fastNlMeansDenoising(enhanced, h=30)
        kernel = np.array([[0, -1, 0], [-1, 5,-1], [0, -1, 0]])
        sharpened = cv2.filter2D(denoised, -1, kernel)
        thresh = cv2.adaptiveThreshold(sharpened, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY, 15, 9)
        final_img = cv2.cvtColor(thresh, cv2.COLOR_GRAY2RGB)
        pil_img = Image.fromarray(final_img)
        tess_text = pytesseract.image_to_string(pil_img, config='--oem 3 --psm 6')
        return tess_text


    def data_cleansing(self, text, card_type, content_type=False):
        try:
            condition = False
            cardno = effective_date = expiry_date = ''
            if not condition and card_type == 'type1':
                reader = easyocr.Reader(['en'],  gpu=False)
                result = reader.readtext(self.image_path, detail=0)
                text = "\n".join(result)
                for index,i in enumerate(text.splitlines()):
                    if 'EID#' in i:
                        cardno = i.split(' ')[-1]
                    elif 'Eff. Date' in i:
                        effective_date = parse(text.splitlines()[index+1]).strftime('%Y-%m-%d')
                    elif 'Expiry Date' in i:
                        expiry_date = parse(text.splitlines()[index+1]).strftime('%Y-%m-%d')
                condition = cardno != '' and effective_date != '' and expiry_date != ''
            # if not condition and card_type == 'type1':
            #     ocr = PaddleOCR(use_angle_cls=True, lang='en')
            #     results = ocr.predict(self.image_path)
            #     results = results[0].get('rec_texts','')
            #     for index,i in enumerate(results):
            #         if 'EID#' in i:
            #             cardno = i.split('#')[-1].split('(')[0]
            #         elif 'Eff.Date' in i:
            #             effective_date = parse(results[index+1].replace('：','').strip()).strftime('%Y-%m-%d')
            #         elif 'Expiry Date' in i:
            #             expiry_date = parse(results[index+1].replace('：','').strip()).strftime('%Y-%m-%d')
            #     condition = cardno != '' and effective_date != '' and expiry_date != ''

            elif not condition and card_type == 'type2':
                for row in text.split('\n\n'):
                    if 'CardNo' in row:
                        cardno = row.split(':')[-1].strip()
                    elif 'Validity' in row:
                        effective_date = parse(row.split(' ')[2]).strftime('%Y-%m-%d')
                        expiry_date = parse(row.split(' ')[-1]).strftime('%Y-%m-%d')
                condition = cardno != '' and effective_date != '' and expiry_date != ''
            
            elif not condition and card_type == 'type3':
                reader = easyocr.Reader(['en'],  gpu=False)
                result = reader.readtext(self.image_path, detail=0)
                text = "\n".join(result)
                for index,i in enumerate(text.splitlines()):
                    if 'Card No' in i:
                        cardno = i.split(':')[-1].strip()
                    elif 'Valid Till' in i:
                        expiry_date = parse(text.splitlines()[index+1]).strftime('%Y-%m-%d')
                    elif 'Valid From' in i:
                        effective_date = parse(text.splitlines()[index+1]).strftime('%Y-%m-%d')

                condition = cardno != '' and effective_date != '' and expiry_date != ''

            elif not condition and card_type == 'type4':
                data = {i.split('\n')[0].replace(' ','').replace('.',''):i.split('\n')[-1] for i in text.split('\n\n')}
                cardno = data.get('CardNo','')
                effective_date = parse(data.get('EffectiveDate','')).strftime('%Y-%m-%d')
                expiry_date = parse(data.get('ExpireDate','')).strftime('%Y-%m-%d')
                condition = cardno != '' and effective_date != '' and expiry_date != ''

            elif not condition and card_type == 'type5':
                reader = easyocr.Reader(['en'],  gpu=False)
                result = reader.readtext(self.image_path, detail=0)
                text = "\n".join(result)
                for index,i in enumerate(text.splitlines()):
                    if 'Card No' in i:
                        cardno = text.splitlines()[index+1]
                    elif 'Card Validity' in i:
                        dates = text.splitlines()[index+1]
                        effective_date = parse(dates.split(' ')[0]).strftime('%Y-%m-%d')
                        expiry_date = parse(dates.split(' ')[-1]).strftime('%Y-%m-%d')
                condition = cardno != '' and effective_date != '' and expiry_date != ''
            
            fields = {'CardNo': cardno, 'EffectiveDate': effective_date, 'ExpiryDate': expiry_date}
            condition = cardno != '' and effective_date != '' and expiry_date != ''
            if condition:
                return fields
            else:
                return {}
        except Exception as e:
            print(f"Error Occured: {e}")
            return {}
    
    def detect_card_type(self, text):
        if 'ADNIC' in text:
            self.message_body = '''Dear Employee,

            Kindly refer to attached files for your Table of Benefits and Network List of the Medical Insurance Policy that you are currently enrolled in. Note that ADNIC is cardless, hence you may use your Emirates ID to avail the medical treatment.

            Please note that the network list is subject to change at times, hence we highly advise to view the updated Network List via the App to ensure that you visit an In-Network facility. Outside network treatments will only be reimbursed by the insurance company depending on the claims settlement of your policy (Refer to the TOB)

            Also note that you are eligible for TruDoc Teleconsultation benefit. Please refer to the attached flyer for a short overview on the features and benefits.

            Regards
            Cards Team
            '''
            return 'type1'
        elif 'TASC OUTSOURCING' in text:
            self.message_body = '''Dear Employee,

            Kindly refer to attached files for your Table of Benefits and Network List of the Medical Insurance Policy that you are currently enrolled in. Note that QIC is cardless, hence you may use your Emirates ID to avail the medical treatment.

            Please note that the network list is subject to change at times, hence we highly advise to view the updated Network List via the App to ensure that you visit an In-Network facility. Outside network treatments will only be reimbursed by the insurance company depending on the claims settlement of your policy (Refer to the TOB)

            Also note that you are eligible for TruDoc Teleconsultation benefit. Please refer to the attached flyer for a short overview on the features and benefits.            

            Regards
            Cards Team
            '''
            return 'type2'
        elif 'HEALTHCARE' in text:
            self.message_body = ''
            return 'type3'
        elif 'Select Silver' in text:
            self.message_body = ''
            return 'type4'
        elif 'MetLife' in text:
            self.message_body = '''Dear Employee,
            Kindly refer to attached files for the e-medical card, Table of Benefits (TOB) and Network List of the policy that you are enrolled in. Note that Metlife is cardless, hence you may use your Emirates ID to avail the medical treatment.

            Please note that the network list is subject to change at times, hence we highly advise to view the updated Network List via the App/portal to ensure that you visit an In-Network facility. Outside network treatments will only be reimbursed by the insurance company depending on the claims settlement of your policy (Refer to the TOB)

            Regards
            Cards Team

            '''
            return 'type5'
        else:
            return None


    def extract_card_information(self, code):
        try:
            jpg_card = [i for i in os.listdir(self.card_dir) if (i.startswith(str(code))) and (i.endswith('.jpg') or i.endswith('.png')) ]
            for img_path in jpg_card:
                self.image_path = os.path.join(self.card_dir, img_path)
                img = Image.open(self.image_path)
                text = pytesseract.image_to_string(img)
                card_type = self.detect_card_type(text)
                image_resolution = self.get_image_resolution()
                # data = self.data_cleansing(text, card_type)
                # data['image_resolution'] = image_resolution
                self.tesseract_text()
                data = []
                return data
                    
        except Exception as e:
            print(f"Error Occured {e}")
            return []

    def insurance_email_draft(self, sender_email, sender_password, recipient_email, subject, body, attachment_dir, data):
        self.email_status = False
        msg = EmailMessage()
        msg["From"] = sender_email
        msg["To"] = recipient_email
        msg["Subject"] = subject
        msg.set_content(body)

        
        try:
            attachment_paths = [os.path.join(attachment_dir, i) for i in os.listdir(attachment_dir) if (i.startswith(str(data['EmpCode']))) and (i.endswith('jpg') or i.endswith('png'))]
            root_dir = r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\ApplicantMaster\GENERAL'
            directory = [i for i in os.listdir(root_dir) if data['MedicalInsurance'] == i]
            if directory:
                insurance_dir = directory[0] if directory else []
                main_dir = os.path.join(root_dir, insurance_dir)
                insurance_docs = [os.path.join(main_dir,i) for i in os.listdir(main_dir)]
                final_docs = attachment_paths+insurance_docs
                for attachment_path in final_docs:
                    with open(attachment_path, "rb") as f:
                        file_data = f.read()
                        file_name = os.path.basename(attachment_path)
                        msg.add_attachment(
                            file_data,
                            maintype="application",
                            subtype="vnd.openxmlformats-officedocument.wordprocessingml.document",
                            filename=file_name,
                        )

            
                with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                    server.login(sender_email, sender_password)
                    server.send_message(msg)
                    self.email_status = True
                print("Email sent successfully.")
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

            url = "https://api22preview.sapsf.com/odata/v2/upsert?"
            username = "SFAPI@tasclabourT1"
            password = "Exalogic@123456"

            payload = {
                "__metadata": {"uri": "Attachment"},
                "fileName": os.path.basename(self.img_file_path),
                "module": "GENERIC_OBJECT",
                "userId": "10900132",
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
                if key_tag:
                    attachment_id = key_tag.text.split("=")[-1]
                    doc_payload  = {
                        "__metadata": {
                            "uri": "cust_DocIDInfowExpiry",
                            "type": "SFOData.cust_DocIDInfowExpiry"
                        },
                        "effectiveStartDate": f"/Date({self.convert_dt_milliseconds(final_data['EffectiveDate'])})/",
                        "externalCode": "10900132",
                        "cust_DocDetails": {
                            "cust_DocumentNumber":final_data['CardNo'],
                            "cust_Document_Type": "Ec_ARE_53",
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
            if data['UploadStatus'] == 'Success':
                shutil.move(self.img_file_path, r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\ApplicantMaster\Done')
            else:
                shutil.move(self.img_file_path, r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\ApplicantMaster\Error')
        except Exception as e:
            print(f"Error Occured: {e}")

    def collect_emp_info(self):
        df = pd.read_csv(self.file_path)
        employee_code = [int(str(i).split('.')[0]) for i in os.listdir(self.card_dir) if str(i).split('.')[0].isdigit()]
        content_list = []
        for code in employee_code:
            df_filter = df[df['ApplicantId'] == code]
            final_data = {'MedicalInsurance':'','CandidateEmailID':'', 'CardNo':'', 'ExpiryDate':'', 'EffectiveDate':'', 'UploadStatus':'', 'ExtractionStatus':'','EmailStatus':'N', 'TASCStatus': 'Completed'}
            final_data['EmpCode'] = code
            
            record2 = self.extract_card_information(code)
            if record2:
                final_data = {**final_data, **record2}
                final_data['ExtractionStatus'] = 'Success'
                # if final_data['ExtractionStatus'] == 'Success':
                #     self.upload_api_data(final_data)
                #     if final_data['UploadStatus'] == 'Success':
                #         self.move_img_file_dir(final_data)
            else:
                final_data['ExtractionStatus'] = 'Failed'

            # if not df_filter.empty:
            #     df_filter = df_filter.rename(columns={'ApplicantId': 'EmpCode', 'Medical Insurance':'MedicalInsurance', 'Candidate Email ID':'CandidateEmailID'})
            #     row = df_filter[['EmpCode', 'MedicalInsurance', 'CandidateEmailID']].to_dict(orient='records')[0]
            #     if '/' in row['MedicalInsurance']:
            #         final_data['MedicalInsurance'] = row['MedicalInsurance'].split('/')[2]
            #         final_data['CandidateEmailID'] = row['CandidateEmailID']
            #         sender_email = "nandhini2571999@gmail.com"  
            #         sender_password = "utfy ogzm qwks vfwt"  
            #         recipient_email = "nandylokesh123@gmail.com" #final_data['CandidateEmailID']  
            #         subject = "Applican Master"
            #         body = "Please find the updated document attached."
            #         attachment_dir = self.card_dir
            #         self.insurance_email_draft(sender_email, sender_password, recipient_email, subject, self.message_body, attachment_dir, final_data)
            #         final_data['EmailStatus'] = 'Y' if self.email_status == True else 'N'
            #     else:
            #         print(f"There is no Medical Insurance Details")
            #         final_data['TASCStatus'] = 'InsuranceDetail Not in the ApplicantMaster'
            # else:
            #     print(f"ApplicanID: {code} Not Found in Excel")
            #     final_data['TASCStatus'] = 'ApplicantID Not in the ApplicanMaster'
            
            content_list.append(final_data) 
        self.generate_logfile(content_list)
        print("Log Generated Successfully")  
    
    def generate_logfile(self, records):
        df = pd.DataFrame(records)
        df = df[['image_resolution','EmpCode', 'CandidateEmailID', 'CardNo', 'ExpiryDate','EffectiveDate','MedicalInsurance', 'EmailStatus', 'ExtractionStatus' , 'UploadStatus' , 'TASCStatus']]
        timestamp = (datetime.now()).strftime('%Y%m%d%H%M%S')

        df.to_excel(f"C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\ApplicantMaster\\Log_{timestamp}.xlsx", index=False)

if __name__ == "__main__":
    app = ApplicantMaster()
    app.collect_emp_info()


r'''
pip install easyocr
pip install pytesseract pillow opencv-python
pip install paddleocr
pip install paddlepaddle

Steps:

Press Start

Search "cmd"

Right-click → Run as administrator

Then install:
pip install opencv-python --upgrade

 3. Manually Delete the Locking File
You can try:

Go to:

vbnet

C:\Users\nandh\MyProject\PythonProjects\pyvenv\Lib\site-packages\cv2
Delete cv2.pyd

Then run:
pip install opencv-python --force-reinstall
'''
