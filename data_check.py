from openpyxl import load_workbook
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module="openpyxl")
import csv


wb = load_workbook(filename='chamber_3.xlsx')

ws = wb['Sheet1 (1)']



#gets list of fax numbers
for i in ws:
    if i[1].value:
        if i[1].value.strip() == "Fax:":
            print(i[1].offset(column=1).value)
#gets list of phone numbers
# for i in ws:
#     if i[1].value:
#         if i[1].value.strip() == "Phone:":
#             print(i[1].value, i[1].offset(column=1).value)
# Gets the email
# for i in ws:
#     if i[1].value:
#         if i[1].value.strip() == "Fax:":
#             print(i[1].value, i[1].offset(column=-1).value)
