import pandas as pd
import json
import glob


try:
    excel_files = [ file for file in glob.glob('C:/Users/ASUS/Desktop/fieldmobi/Upload Folder/*.xlsx') if 'Error_Template' not in file]
    print(f'Excel files found: {excel_files}')

    if excel_files:
        for file in excel_files:
            df = pd.read_excel(file, engine = 'openpyxl', nrows=2)
            data = df.iloc[0, [2, 4]]
            data_str = str(data.iloc[0]) + '_' + str(data.iloc[1])
            print(data_str)

            df = pd.read_excel(file, engine= 'openpyxl', skiprows = 3)
            df = df.drop(columns = ['Sr No'])
            df = df.dropna(how = 'all')

            df = df.rename(columns={
                'Field Name' : 'data',
                'Default Label': 'label',
                'List Type': 'list_type',
                'Default List Value': 'list_value',
                
            })

            if df['data'].isnull().any():
                raise ValueError("Missing value in 'data' column.")
            df = df.drop(columns=['Field Type'])
            print(df)



            json_obj = {}
            for i, row in df.iterrows():
                row_data = {}
                for col in df.columns:
                    if pd.notna(row[col]):
                        row_data[col] = row[col]
                
                
                json_obj['fieldCode' + str(i+1).zfill(3)] = row_data
            json_str = json.dumps(json_obj, indent = 4)

            with open('output_fieldlist.txt', 'w') as file:
                file.write(data_str + '\n\n')
                file.write(json_str)

            print(json_str)
        else:
            print("No Excel files found in the local directory")
except Exception as e:
    print(f"An error occured: {e}")
