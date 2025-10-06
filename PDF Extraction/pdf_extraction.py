import camelot
import pandas as pd
import re
from PyPDF2 import PdfReader

def combine_two_excel_file():
    df1 = pd.read_excel(pdf_filepath.replace('pdf','xlsx'))
    df2 = pd.read_excel(pdf_filepath.replace('pdf','xlsx'))
    df3 = pd.read_excel(r'C:\Users\nandh\MyProject\PythonProjects\TASCOutsourcing\PDF Extraction\JBI -WPS - May  2025.xlsx', skiprows=2)

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

file_paths = ['JBI PROPERTIES SERVICES L.L.C-employee list.pdf','JBI PROPERTIES SERVICES L.L.C-WPS.pdf']
for module in range(2):
    file = file_paths[module]
    if module == 0:
        # Helper function to clean illegal characters
        def remove_illegal_chars(value):
            if isinstance(value, str):
                # Remove characters not allowed in Excel (ASCII < 32 except \t, \n, \r)
                return re.sub(r"[\x00-\x08\x0B-\x0C\x0E-\x1F]", " ", value)
            return value

        pdf_filepath = f'C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\PDF Extraction\{file}'
        # Extract tables
        reader = PdfReader(pdf_filepath)
        total_pages = len(reader.pages)

        tables = camelot.read_pdf(pdf_filepath, pages=f'1-{total_pages}', flavor='lattice')  # 'stream' or 'lattice'
        print(f"Found {len(tables)} table(s)")
        all_tables = pd.concat([table.df for table in tables], ignore_index=True)
        df = all_tables[1:]
        content_list = []
        for index, row in df.iterrows():
            fields = {}
            if 'Person Name' not in str(row[1]):
                try:
                    fields['PassportNumber'] =  row[0].splitlines()[0]
                except:
                    fields['PassportNumber'] =  row[0]
                try:
                    fields['PersonName'] =  row[1].splitlines()[0]
                except:
                    fields['PersonName'] = row[1]
                try:
                    fields['ArabicPersonName'] =  row[1].splitlines()[1]
                except:
                    fields['ArabicPersonName'] = row[1]
                try:
                    fields['PersonCode'] =  row[1].splitlines()[-1]
                except:
                    fields['PersonCode'] = row[1]
                try:
                    fields['CardType'] =  row[2].splitlines()[0]
                except:
                    fields['CardType'] = row[2]
                try:
                    fields['ArabicCardType'] =  row[2].splitlines()[-1]
                except:
                    fields['ArabicCardType'] = row[2]
                try:
                    fields['JobName'] =  row[3].splitlines()[0]
                except:
                    fields['JobName'] = row[3]
                try:
                    fields['ArabicJobName'] =  row[3].splitlines()[-1]
                except:
                    fields['ArabicJobName'] = row[3]
                try:
                    fields['Nationality'] =  row[4].splitlines()[0]
                except:
                    fields['Nationality'] = row[4]
                try:
                    fields['ArabicNationality'] =  row[4].splitlines()[-1]
                except:
                    fields['ArabicNationality'] = row[4]
                try:
                    fields['CardNumber'] =  row[5].splitlines()[0]
                except:
                    fields['CardNumber'] = row[5]
                try:
                    fields['CardDate'] = row[5].splitlines()[-1]
                except:
                    fields['CardDate'] = row[5]
                try:
                    fields['ContractType'] = row[6].splitlines()[0]
                except:
                    fields['ContractType'] = row[6]
                try:
                    fields['ArabicContractType'] = row[6].splitlines()[-1]
                except:
                    fields['ArabicContractType'] = row[6]

                content_list.append(fields)
        df_cleaned = pd.DataFrame(content_list).applymap(remove_illegal_chars)
        df_cleaned.to_excel(pdf_filepath.replace('pdf','xlsx'), index=False)
        print("PDF To Excel Converted Successfully!")

    elif module ==1:
        pdf_filepath = f'C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\PDF Extraction\{file}'
        # Extract tables
        reader = PdfReader(pdf_filepath)
        total_pages = len(reader.pages)
        tables = camelot.read_pdf(pdf_filepath, pages=f'1-{total_pages}', flavor='lattice')  # 'stream' or 'lattice'
        print(f"Found {len(tables)} table(s)")
        all_tables = pd.concat([table.df for table in tables], ignore_index=True)
        df = all_tables.loc[6:]
        content_list = []
        counter = 0
        for index, row in df.iterrows():
            try:
                if [i for i in row[0]][0] in ['1','2','3','4','5','6','7','8','9','0']:
                    fields = {
                    'RowNumber': row[0],
                    'ArabicPersonName': row[1].splitlines()[0],
                    'PersonName': row[1].splitlines()[1],
                    'PersonCode': row[1].splitlines()[2],
                    'Paid': row[2].splitlines()[0],
                    'Contract': (row[2].splitlines()[1]).split(':')[-1].strip(),
                    'WPSFullfillment': 'Tick - Yes' if row[3] == '\uf00c' else row[3]
                    }
                    if row[0] == '111':
                        if len(row[2].splitlines()) == 4:
                            fields['Contract'] = row[2].splitlines()[2]
                            fields['Paid'] = row[2].splitlines()[1]
                            fields['WPSFullfillment'] = row[2].splitlines()[0]+row[2].splitlines()[-1]
                    content_list.append(fields)
            except Exception as e:
                print(f"Error {e}", row)
        final_result = pd.DataFrame(content_list).to_excel(pdf_filepath.replace('pdf','xlsx'),index = False)
combine_two_excel_file()