import pandas as pd
import json

df = pd.read_excel('Copy of First ERP Configurations_ PEOPLE (1).xlsx', engine='openpyxl', nrows=2)
data = df.iloc[0, [2, 4]]
data_str = str(data.iloc[0]) + '_' + str(data.iloc[1])
print(data_str)


# Load the Excel file into a DataFrame
df = pd.read_excel('Copy of First ERP Configurations_ PEOPLE (1).xlsx', engine='openpyxl', skiprows=3)

# Remove the rows where all the elements are missing

df = df.drop(columns=['Sr No'])

# Select only the first 46 rows
df = df.iloc[:39]
df = df.iloc[1:]

print(df)

# Initialize an empty dictionary to hold the JSON object
json_obj = {}

# Iterate over the rows of the DataFrame
for i, row in df.iterrows():
    # Initialize an empty dictionary to hold the data for the current row
    row_data = {}
    
    # Iterate over the columns of the DataFrame
    for col in df.columns:
        # Add the data for the current column to the row data
        row_data[col] = row[col]
    
    # Add the row data to the JSON object with a key of 'fieldcode' followed by the row number
    json_obj['fieldCode' + str(i).zfill(3)] = row_data

# Convert the dictionary to a JSON object with an indentation of 4
json_str = json.dumps(json_obj, indent=4)

with open('output.txt', 'w') as file:
    file.write(data_str + '\n\n')
    file.write(json_str)

# Print the JSON object
print(json_str)






