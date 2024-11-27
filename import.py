import openpyxl
import json
import requests

#set pipline id and PipelineStageId  here
piplineid=    
PipelineStageId= 
OwnerId=


def read_excel_until_blank_row(file_path, sheet_name):
    # Load the workbook
    workbook = openpyxl.load_workbook(file_path)

    # Select the specified sheet
    sheet = workbook[sheet_name]

    # Initialize an empty list to store the data
    data = []

    # Iterate through rows
    for row in sheet.iter_rows(values_only=True):
        # Check if the row is blank (all cells are None or empty string)
        if all(cell is None or cell == '' for cell in row):
            break  # Exit the loop when a blank row is found

        # Add non-blank row to the data list
        data.append(row)

    # Close the workbook
    workbook.close()

    return data


def importtodidar (rowdata) :

    url = "https://app.didar.me/api/case/SaveCase"

    payload = json.dumps({
    "Case": {
        "DueDate": "9999-12-01T00:00:00.000Z",
        "Priority": -2,
        "PipelineStageId": PipelineStageId,
        "OwnerId": OwnerId,,
        "SegmentIds": [],
        "Title": rowdata,
        "VisibilityType": "All",
        "Fields": {
        "Field_8785_14_78": "مارکتپلیس"
        },
        "CompanyId": "00000000-0000-0000-0000-000000000000",
        "PersonId": "00000000-0000-0000-0000-000000000000"
    },
    "SetLabelDto": {
        "UpdatedLabels": [],
        "AddedLabels": [],
        "LabelIds": [
        "ebb2a5c5-d295-4ebc-9f61-0082b140d9b2"
        ]
    }
    })
    headers = {
    'accept': 'application/json, text/plain, */*',
    'accept-language': 'en-US,en;q=0.9,fa;q=0.8',
    'authorization': '8e5502ab-1487-416e-ba92-a7fd66b3d390',
    'content-type': 'application/json',
    'priority': 'u=1, i',
    'referer': 'https://app.didar.me/case;pipeline=piplineid    ;_force=true;tab=owner',
    'sec-ch-ua': '"Chromium";v="130", "Google Chrome";v="130", "Not?A_Brand";v="99"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Windows"',
    'sec-fetch-dest': 'empty',
    'sec-fetch-mode': 'cors',
    'sec-fetch-site': 'same-origin',
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36',
    'x-bizdomain': 'armitis',
    'x-url': 'https://app.didar.me/case;pipeline=piplineid;_force=true;tab=owner'
    }

    response = requests.request("POST", url, headers=headers, data=payload)

    print(response.text)

# Example usage
file_path = 'Book1.xlsx'
sheet_name = 'Sheet1'

excel_data = read_excel_until_blank_row(file_path, sheet_name)

# Print the data
for row in excel_data:
    print(row[2])
    importtodidar(row[2])