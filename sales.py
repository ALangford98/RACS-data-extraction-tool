import openpyxl
import pandas as pd

ps = openpyxl.load_workbook('sales.xlsx', data_only=True)
sheet = ps['Feb17']
rng = range(11, 168 + 1)
for row in rng:
    date = str(sheet['C' + str(row)].value).replace(" ", str(None))
    salesAmt = str(sheet['D' + str(row)].value)
    product = str(sheet['E' + str(row)].value)
    salesQty = str(sheet['F' + str(row)].value)
    
    if not date == "None":
        lDate = date
    if date =="None":
        date = lDate
    else:
        date = date
    out = (date + "/02/2017, "+product+", "+ salesAmt+", "+ salesQty)
    print(out)