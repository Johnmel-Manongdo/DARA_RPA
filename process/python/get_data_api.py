import requests
from requests_toolbelt import MultipartEncoder
import os
import json
import pandas as pd
import openpyxl

app_id = "cli_a2bffebdbd78d009"
app_secret = "hQBykhX6T8iLGmzzJGHzxh0rBE0iBAOF"
app_token = "basusm83Gkfxi4Uld8BJSQHUFkN"
table_id = "tbl1AUpJuXiDL12T"
view_id = "vewp7nmiS4"

def process():

    ## Request Token
    request_token_url = "https://open.larksuite.com/open-apis/auth/v3/tenant_access_token/internal"
    request_token_data = {
            "app_id": app_id,
            "app_secret": app_secret
            }
    gettoken = requests.post(url=request_token_url, data=request_token_data)
    tenant_access_token = gettoken.json()['tenant_access_token']

    ## Request Records
    url = "https://open.larksuite.com/open-apis/bitable/v1/apps/"+app_token+"/tables/"+table_id+"/records"
    headers = {
    'Authorization': 'Bearer '+tenant_access_token
    }
    params = {}
    response = requests.get(url, headers=headers, params=params)

    if response.status_code == 200:

        all_records = response.json()['data']
        # print(all_records['items'])

        list = []
        for item in all_records['items']:

            list.append(item['fields'])
            if "STATUS" in item['fields']:

                if(item['fields']['STATUS'] == 'APPROVED'):

                    print(item['fields']['STATUS'])

                    url = "https://open.larksuite.com/open-apis/bitable/v1/apps/"+app_token+"/tables/"+table_id+"/records/"+item['record_id']
                    payload = json.dumps({
                            "fields": {
                                "REMARKS": "Done"
                            },
                            "record_id": item['record_id']
                        })

                    headers = {
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer '+tenant_access_token
                    }

                    requests.put(url, headers=headers, data=payload)

            else:
                print("Not exist")

        df = pd.DataFrame(list)
        df['CH NAME'] = df['CH NAME'].str.strip() # remove extra spaces at the end of CH NAME column
        filtered_df = df.loc[(df['STATUS'] == 'APPROVED') & (df['STATUS'] != 'DISAPPROVED') & (df['REMARKS'] != 'Done'), ~(df.columns.str.contains('APPROVER|CREATED BY'))]

        ## Get column names to format
        columns_to_format = ['ENDORSED OR CURRENT OB', '2ND ENDORSED OR CURRENT OB', '3RD ENDORSED OR CURRENT OB', 'MONTHLY IF:', '2ND MONTHLY IF:', '3RD MONTHLY IF:', 'PAYMENT AMOUNT', '2ND PAYMENT AMOUNT', '3RD PAYMENT AMOUNT', 'TOTAL PAYABLE AMOUNT']
        currency_columns = ['MONTHLY IF:', '2ND MONTHLY IF:', '3RD MONTHLY IF:', 'TOTAL PAYABLE AMOUNT']

        ## Format the data
        for column in columns_to_format:
            if column in filtered_df.columns:
                if column in currency_columns:
                    filtered_df[column] = pd.to_numeric(filtered_df[column], errors='coerce')
                    filtered_df[column] = filtered_df[column].apply(lambda x: '₱ {:,.2f}'.format(x) if pd.notnull(x) else '')
                else:
                    filtered_df[column] = pd.to_numeric(filtered_df[column], errors='coerce')
                    filtered_df[column] = filtered_df[column].apply(lambda x: '{:,.2f}'.format(x) if pd.notnull(x) else '')

        columns_to_format_percent = ['DISCOUNT', 'INTEREST RATE', '2ND DISCOUNT', '3RD DISCOUNT', '2ND INTEREST RATE', '3RD INTEREST RATE']

        for column_percent in columns_to_format_percent:
            if column_percent in filtered_df.columns:
                filtered_df[column_percent] = pd.to_numeric(filtered_df[column_percent], errors='coerce')
                filtered_df[column_percent] = filtered_df[column_percent].apply(lambda x: '{:.0%}'.format(x) if pd.notnull(x) else '')
        
##       # Write the DataFrame to an Excel file
##        with pd.ExcelWriter('DARA Request.xlsx') as writer:
##            filtered_df.to_excel(writer, sheet_name='Sheet1', index=False)

         ## Get the directory of config file
        output_file = os.path.join(os.path.expanduser("~/Desktop"), "DARA BOT", "config", "DARA Generator", "config.xlsm")

        ## Load the existing Excel file
        workbook = openpyxl.load_workbook(output_file, keep_vba=True)

        ## Select the sheet to write to
        worksheet = workbook['List']

        ## Iterate over each column in the worksheet
        for col in worksheet.iter_cols():
            header = col[0].value
            if header in filtered_df.columns:
                ## Get the index of the column in the DataFrame
                col_index = filtered_df.columns.get_loc(header)
                ## Get the data for that column
                col_data = filtered_df.iloc[:, col_index]
                ## Write the data to the corresponding cells in the worksheet
                for cell_index, cell_value in enumerate(col_data):
                    cell = col[cell_index+1]
                    cell.value = cell_value

        ## Save the changes to the Excel file
        workbook.save(output_file)

        num_rows = len(filtered_df)
        print(num_rows)

process()
