import pandas as pd
import json
import glob

try:
    excel_files = [ file for file in glob.glob('C:/Users/ASUS/Desktop/fieldmobi/Upload Folder/*.xlsx') if 'Error_Template' not in file]
    print(f'Excel files found: {excel_files}')

    if excel_files:
        for file in excel_files:
            df = pd.read_excel(file, engine='openpyxl', nrows = 3)
            print(df)
            data1 = df.iloc[0,2]
            data2 = df.iloc[1,2]
            data3 = df.iloc[1,4]
            data_str = str(data1) + '_' + str(data2) + '_' + str(data3)
            print(data_str)

            df = pd.read_excel(file, engine='openpyxl', skiprows=3)
            df = df.drop(columns=['Sr No'])
            df = df.dropna(how = 'all')

            df.reset_index(drop = True, inplace = True)

            df = df.rename(columns={
                'Field Name' : 'data',
                'New Label': 'label'
                
            })

            if df['data'].isnull().any():
                raise ValueError("Missing value in 'data' column.")

            df.rename(columns = {
                'Mobile Seq': 'Mobile_Mobile Seq',
                'Validation': 'Mobile_Validation',
                'Link Setup': 'Mobile_Link Setup',
                'Update Setup': 'Mobile_Update Setup'
            }, inplace = True)

            df.rename(columns={
                'Search Seq': 'Web_Search Seq',
                'Display Seq': 'Web_Display Seq',
                'Create Seq': 'Web_Create Seq',
                'Edit Seq': 'Web_Edit Seq',
                'Validation.1': 'Web_Validation',
                'Link Setup.1': 'Web_Link Setup',
                'Update Setup.1': 'Web_Update Setup',
                'List Seq': 'Web_List Seq',
                'Summary Seq': 'Web_Summary Seq',
                'Map Seq': 'Web_Map Seq',
                'Report Seq': 'Web_Report Seq'
            }, inplace=True)
    

            json_obj = {}

            for i, row in df.iterrows():
                row_data = {}
                for col in df.columns:
                    if pd.notna(row[col]):         
                        row_data[col] = row[col]
                json_obj['fieldCode' + str(i+1).zfill(3)] = row_data

            json_str = json.dumps(json_obj, indent=4)

            with open('output_dataview.txt', 'w') as file:
                file.write(data_str + '\n\n')
                file.write(json_str)

            print(json_str)

        else:
            print("No Excel Files found in the local directory")
except Exception as e:
    print(f"An error occured: {e}")



