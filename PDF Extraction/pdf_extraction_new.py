import pdfplumber
import pandas as pd
import re

def combine_two_excel_file(file1,file2, exccel):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    df3 = pd.read_excel(exccel, skiprows=2)

    df1['PersonCode'] = df1['PersonCode'].astype(str)
    df2['PersonCode'] = df2['PersonCode'].astype(str)
    df3['PersonCode'] = df3['PersonCode'].astype(str)
    data_df = pd.merge(df1, df2, on='PersonCode', how='left')
    # ['PassportNumber', 'PersonName', 'ArabicPersonName_x', 'PersonCode', 'CardType', 'ArabicCardType', 'JobName', 'ArabicJobName', 'Nationality', 'ArabicNationality', 'CardNumber', 'CardDate', 'ContractType', 'ArabicContractType', 'RowNumber', 'ArabicPersonName_y', 'Paid', 'Contract', 'WPSFullfillment']
    data_df = data_df[['PersonCode', 'PersonName_x', 'JobName', 'PassportNumber', 'Nationality', 'CardNumber', 'CardType', 'CardDate', 'Contract']]
    data_df = data_df.merge(df3[['PersonCode', 'WBSNumber', 'SAPNumber']], on='PersonCode', how='left')
    data_df = data_df.rename(columns={'PersonName_x': 'PersonName', 'JobName':'Job'})
    data_df.to_excel(r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\PDF Extraction\pdf_extraction_output.xlsx', index = False)
    print("Successfully Generated")

# Helper function to clean illegal characters
def remove_illegal_chars(value):
    if isinstance(value, str):
        # Remove characters not allowed in Excel (ASCII < 32 except \t, \n, \r)
        return re.sub(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]", " ", value)
    return value

def get_pdf_tables(filename):
    all_rows = []
    with pdfplumber.open(filename) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()
            all_rows.append(tables)
    return all_rows

excel_file = f'C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\PDF Extraction\\Serco Limited Dubai Branch - WPS - APR - 2025.xlsx'
input_pdf1  = f'C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\PDF Extraction\\JBI PROPERTIES SERVICES L.L.C-employee list.pdf'
all_rows = get_pdf_tables(input_pdf1)
passport_details = [j for i in all_rows for j in i if 'Passport Number' in str(j)]
pdf_content = [j for i in passport_details for j in i]

content_list = []
for row in pdf_content:
    if row[0] != '':
        if 'Person Name' not in str(row[1]):
            try:
                fields = {
                'PassportNumber': row[0],
                'PersonName': row[1].splitlines()[0],
                'ArabicPersonName': row[1].splitlines()[1],
                'PersonCode': row[1].splitlines()[-1],
                'CardType': row[2].splitlines()[0],
                'ArabicCardType': row[2].splitlines()[-1],
                'JobName': row[3].splitlines()[0],
                'ArabicJobName': row[3].splitlines()[-1],
                'Nationality': row[4].splitlines()[0],
                'ArabicNationality': row[4].splitlines()[-1],
                'CardNumber': row[5].splitlines()[0],
                'CardDate': row[5].splitlines()[-1],
                'ContractType': row[6].splitlines()[0],
                'ArabicContractType': row[6].splitlines()[-1]
                }
                content_list.append(fields)
            except:
                print(row)
        else:
            print(row)
    else:
        if str(row[1]).isdigit():
            fields.update({'PersonCode': row[1]})
        else:
            print(row)
df_cleaned = pd.DataFrame(content_list).applymap(remove_illegal_chars)
df_cleaned.to_excel(input_pdf1.replace('.pdf','.xlsx'), index=False)
print("PDF To Excel Converted Successfully!")

input_pdf2 = f'C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\PDF Extraction\\JBI PROPERTIES SERVICES L.L.C-WPS.pdf'
all_rows = get_pdf_tables(input_pdf2)

content_list = []
pdf_content = [k for i in all_rows for j in i for k in j if str(k[0]).isdigit()]
for row in pdf_content:
    fields = {
    'RowNumber': row[0],
    'ArabicPersonName': row[1].splitlines()[0],
    'PersonName': row[1].splitlines()[1],
    'PersonCode': row[1].splitlines()[2],
    'Paid': row[2].splitlines()[0],
    'Contract': (row[2].splitlines()[1]).split(':')[-1].strip(),
    'WPSFullfillment': 'Tick - Yes' if row[3] == '\uf00c' else row[3]
    }
    content_list.append(fields)

df_cleaned = pd.DataFrame(content_list).applymap(remove_illegal_chars)
df_cleaned.to_excel(input_pdf2.replace('.pdf','.xlsx'), index=False)
print("PDF To Excel Converted Successfully!")


combine_two_excel_file(input_pdf1.replace('.pdf','.xlsx'), input_pdf2.replace('.pdf','.xlsx'), excel_file)