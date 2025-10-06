import pdfplumber
import pandas as pd
import re
import os
import shutil

def combine_two_excel_file(file1,file2, exccel, output_folder):
    df1 = pd.read_excel(file1)
    df2 = pd.read_excel(file2)
    df3 = pd.read_excel(exccel, skiprows=2)

    df1['PersonCode'] = df1['PersonCode'].astype(str)
    df2['PersonCode'] = df2['PersonCode'].astype(str)
    df3['PersonCode'] = df3['PersonCode'].astype(str)
    data_df = pd.merge(df1, df2, on='PersonCode', how='left')
    # ['PassportNumber', 'PersonName', 'ArabicPersonName_x', 'PersonCode', 'CardType', 'ArabicCardType', 'JobName', 'ArabicJobName', 'Nationality', 'ArabicNationality', 'CardNumber', 'CardDate', 'ContractType', 'ArabicContractType', 'RowNumber', 'ArabicPersonName_y', 'Paid', 'Contract', 'WPSFullfillment']
    data_df = data_df[['PersonCode', 'PersonName_x', 'JobName', 'PassportNumber', 'Nationality', 'CardNumber', 'CardType', 'CardDate', 'Contract', 'Paid']]
    data_df = data_df.merge(df3[['PersonCode', 'WBSNumber', 'SAPNumber']], on='PersonCode', how='left')
    data_df = data_df.rename(columns={'PersonName_x': 'PersonName', 'JobName':'Job'})
    data_df.to_excel(f'{output_folder}\\pdf_extraction_output.xlsx', index = False)
    print("Successfully Generated")
    return True
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

def remove_all_files_in_folder(folder_path):
    if not os.path.isdir(folder_path):
        print(f"{folder_path} is not a valid directory.")
        return

    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)  # remove file or symbolic link
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)  # remove sub-directory
        except Exception as e:
            print(f"Failed to delete {file_path}. Reason: {e}")
            
def main():
    input_folder = 'C:\\Users\\BotServer_2.WIN-IJ6UPG7E58S\\OneDrive - TASC LABOUR SERVICES L.L.C\\PRO_Outsourcing - TCS AUTOMATIONS\\PDF_Extraction\\Input'
    input_folder = 'C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\PDF Extraction\\Input'#
    master_folder = 'C:\\Users\\BotServer_2.WIN-IJ6UPG7E58S\\OneDrive - TASC LABOUR SERVICES L.L.C\\PRO_Outsourcing - TCS AUTOMATIONS\\PDF_Extraction\\Master_Data'
    master_folder = 'C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\PDF Extraction\\Master_Data'#
    
    filenames = os.listdir(input_folder)
    try:
        meta_data = {'Input1': filenames[0],
                 'Input2': filenames[-1],
                 'Output1': filenames[0].replace('.pdf','.xlsx'),
                 'Output2': filenames[-1].replace('.pdf','.xlsx'),
                 'Master_File': [i for i in os.listdir(master_folder) if filenames[0].startswith(i[0:10])][0],
                }
        Master_file = f'C:\\Users\\BotServer_2.WIN-IJ6UPG7E58S\\OneDrive - TASC LABOUR SERVICES L.L.C\\PRO_Outsourcing - TCS AUTOMATIONS\\PDF_Extraction\\Master_Data\\{meta_data["Master_File"]}'
        Master_file = f'C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\PDF Extraction\\Master_Data\\{meta_data["Master_File"]}'#
        input_pdf1  = f'C:\\Users\\BotServer_2.WIN-IJ6UPG7E58S\\OneDrive - TASC LABOUR SERVICES L.L.C\\PRO_Outsourcing - TCS AUTOMATIONS\\PDF_Extraction\\Input\\{meta_data["Input1"]}'
        input_pdf1  = f'C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\PDF Extraction\\Input\\{meta_data["Input1"]}'#
        input_pdf2 = f'C:\\Users\\BotServer_2.WIN-IJ6UPG7E58S\\OneDrive - TASC LABOUR SERVICES L.L.C\\PRO_Outsourcing - TCS AUTOMATIONS\\PDF_Extraction\\Input\\{meta_data["Input2"]}'
        input_pdf2 = f'C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\PDF Extraction\\Input\\{meta_data["Input2"]}'#
        Excel_file1 = f'C:\\Users\\BotServer_2.WIN-IJ6UPG7E58S\\OneDrive - TASC LABOUR SERVICES L.L.C\\PRO_Outsourcing - TCS AUTOMATIONS\\PDF_Extraction\\Output\\{meta_data["Output1"]}'
        Excel_file1 = f'C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\PDF Extraction\\Output\\{meta_data["Output1"]}'#
        Excel_file2 = f'C:\\Users\\BotServer_2.WIN-IJ6UPG7E58S\\OneDrive - TASC LABOUR SERVICES L.L.C\\PRO_Outsourcing - TCS AUTOMATIONS\\PDF_Extraction\\Output\\{meta_data["Output2"]}'
        Excel_file2 = f'C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\PDF Extraction\\Output\\{meta_data["Output2"]}'#
        Completed_folder = 'C:\\Users\\BotServer_2.WIN-IJ6UPG7E58S\\OneDrive - TASC LABOUR SERVICES L.L.C\\PRO_Outsourcing - TCS AUTOMATIONS\\PDF_Extraction\\Completed'
        Completed_folder = 'C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\PDF Extraction\\Completed'#
        Output_folder = 'C:\\Users\\BotServer_2.WIN-IJ6UPG7E58S\\OneDrive - TASC LABOUR SERVICES L.L.C\\PRO_Outsourcing - TCS AUTOMATIONS\\PDF_Extraction\\Output'
        Output_folder = 'C:\\Users\\nandh\\MyProject\\PythonProjects\\TASCOutsourcing\\PDF Extraction\\Output'#
        # Check if at least one PDF file exists
        if not os.path.exists(input_pdf1) and not os.path.exists(input_pdf2):
            print("No PDF files found. Exiting.")
            return
        remove_all_files_in_folder(Completed_folder)
        remove_all_files_in_folder(Output_folder)
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
        df_cleaned.to_excel(Excel_file1, index=False)
        print("PDF To Excel Converted Successfully!")
        all_rows = get_pdf_tables(input_pdf2)

        content_list = []
        pdf_content = [k for i in all_rows for j in i for k in j if str(k[0]).isdigit()]
        for row in pdf_content:
            fields = {
            'RowNumber': row[0],
            'ArabicPersonName': row[1].splitlines()[0],
            'PersonName': row[1].splitlines()[1],
            'PersonCode': row[1].splitlines()[2],
            'Paid': row[2].splitlines()[0].split(':')[-1].strip(),
            'Contract': (row[2].splitlines()[1]).split(':')[-1].strip(),
            'WPSFullfillment': 'Tick - Yes' if row[3] == '\uf00c' else row[3]
            }
            content_list.append(fields)

        df_cleaned = pd.DataFrame(content_list).applymap(remove_illegal_chars)
        df_cleaned.to_excel(Excel_file2, index=False)
        print("PDF To Excel Converted Successfully!")
        
        if os.path.exists(Excel_file1) or os.path.exists(Excel_file2):
            status = combine_two_excel_file(Excel_file1, Excel_file2, Master_file, Output_folder)
            if status == True:
                shutil.move(input_pdf1, Completed_folder)
                shutil.move(input_pdf2, Completed_folder)
                print("Process Completed!")
            else:
                print("Unable to convert the file from pdf to Excel")

    except Exception as e:
        print(f"Error Occured: {e}")  
if __name__ == "__main__":
    main()