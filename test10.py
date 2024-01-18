import pandas as pd
import runpy
import glob

try:
    excel_files = [ file for file in glob.glob('C:/Users/ASUS/Desktop/fieldmobi/Upload Folder/*.xlsx') if 'Error_Template' not in file]
    print(f'Excel files found: {excel_files}')

    if excel_files:
        for file in excel_files:
            print(f"Processing file: {file}")

            df = pd.read_excel(file, engine = 'openpyxl')
            print(df)

            try:
                print(f"Value at H3: {df.iloc[1, 7]}")
                print(f"Value at L3: {df.iloc[1, 11]}")
                print(f"value at A4 : {df.iloc[4,1]}")
                print(f"value at B4 : {df.iloc[4,2]}")

                if df.iloc[1, 7] == 'Mobile Configuration' and df.iloc[1,11] == 'Web Configuration':

                    runpy.run_path('test9.py')
                elif df.iloc[4,1] == 'KEY' and df.iloc[4,2] =='TYP':
                    runpy.run_path('test11.py')
                else:
                    runpy.run_path('test8.py')
            except IndexError:
                print("IndexError: Index is out of bounds for axis, running test8")
                

    else:
        print("No Excel files found in the local directory")
except Exception as e:
    print(f"An error occured: {e}")

