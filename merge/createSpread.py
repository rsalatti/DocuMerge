import auth
import sys
from auth import client
import json

spreadOwner=sys.argv[1]
jsonFile=sys.argv[2]

with open(jsonFile) as f:
        data = json.load(f)
        spreadName = (data['documentName'])
        if ".docx" in spreadName:
                spreadName = spreadName.replace('.docx','')
        elif ".pdf" in spreadName:
                spreadName = spreadName.replace('.pdf',',')

def createSpreadsheet(spreadName,spreadOwner):
        keyCount = 0
        titles_list = []
        for spreadsheet in client.openall():
                titles_list.append(spreadsheet.title)

        if spreadName not in titles_list:
                sh = client.create(spreadName)
                sh.share(spreadOwner, perm_type='user', role='writer')
                sheet = client.open(spreadName).sheet1
                for key,value in data.items():
                        keyCount+=1
                        sheet.update_cell(1,keyCount, key)
                        sheet.update_cell(2,keyCount, value)
        else:
                sheet = client.open(spreadName).sheet1
                next_row = next_available_row(sheet)
                for key, value in data.items():
                        keyCount+=1
                        sheet.update_cell(next_row,keyCount, value)

# Determine first empty row
def next_available_row(sheet):
        str_list = list(filter(None, sheet.col_values(1)))
        return str(len(str_list)+1)

createSpreadsheet(spreadName,spreadOwner)
