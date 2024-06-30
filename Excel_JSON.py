import pandas as pd
from json import dumps
# Load the Excel file
datapath = r"C:\Users\Nicole\Desktop\python_projects\DE Guide 2024\CDP_2018.xlsx"
excel_data = pd.ExcelFile(datapath)

# Initialise a dictionary to store merged data
merged_data = {}

# Define metadata fiels
excl_fields = ['Organization', 'country/area', 'Account number', 'Primary activity', 'Primary sector', 'Primary industry', 'Tickers', 'Authority types', 'Row', 'RowName']
#Iterate through each sheet in the workbook
for sheet_name in excel_data.sheet_names:
    df = pd.read_excel(datapath, sheet_name=sheet_name)
    for index, row in df.iterrows():
        account_number = row['Account number']

        # Metadata subdictionary
        metadata = {
            'Organization' : row['Organization'],
            'Account ID' : row['Account number'],
            'Country': row['country/area'],
            'Primary activity': row['Primary activity'],
            'Primary sector' : row['Primary sector'],
            'Primary industry' : row['Primary industry']
        }

        row_dict = {key : value for key, value in row.to_dict().items() if key not in excl_fields}

        if account_number not in merged_data:
            merged_data[account_number] = {'metadata':metadata}
            merged_data[account_number]['metadata'][sheet_name] = {sheet_name: row_dict}

        
        else:
            # Update exisiting metadata with new values if not NaN
            for key, value in metadata.items():
                if pd.notna(value) and (key not in merged_data[account_number]['metadata'] or
                                        pd.isna(merged_data[account_number]['metadata'][key])):
                    merged_data[account_number]['metadata'][key] = value 
            if 'sheets' not in merged_data[account_number]['metadata']:
                merged_data[account_number]['metadata']['sheets'] = {sheet_name: row_dict}
            elif sheet_name not in merged_data[account_number]['metadata']['sheets']:
                merged_data[account_number]['metadata']['sheets'][sheet_name] = row_dict
            else:
                for key, value in row_dict.items():
                    if pd.notna(value):
                        merged_data[account_number]['metadata']['sheets'][sheet_name][key] = value
     

# Convert the dictionary to a JSON string
json_data = dumps(merged_data, indent=4)
# Optionally, write the JSON data to a file 
with open ('output.json', 'w') as json_file:
    json_file.write(json_data)