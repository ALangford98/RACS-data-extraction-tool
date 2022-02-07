import openpyxl
import pandas as pd

ps = openpyxl.load_workbook('.xlsx') #Spreadsheet file containing raw data
sheet = ps['sheet1'] #sheet to use while extracting data
for row in range(1, 63 +1): #rows to consider for extraction
    surname = sheet['C' + str(row)].value
    arrears = str(sheet['D' + str(row)].value)
    installment = str(sheet['E' + str(row)].value)
    ref = str(sheet['F' + str(row)].value).lstrip("'")
    phone = str(sheet['G' + str(row)].value).rstrip("0.")
    client = sheet['A' + str(row)].value
    ptp = str(sheet['H' + str(row)].value).rstrip("00:00:00")
    #out = "\nDear Mr/Ms "+surname+"\nEnsure to urgently pay Hi-Finance account the arrears or minimum monthly installment or more.\nArrears: N$"+arrears+"\nInstallment: N$"+installment+"\nReference: "+ref+"\nPayment by 30 November 2020 at nearest JDG Store:\nHi-Fi Corporation, Sleepmasters or Incredible Connection\nContact:Resolve account collection services\nTel:061307491 081"+phone+"\n"
    #print(out)
    print("-------------------------------------------------------------------------------------------------")
    print(row)
    print("Dear Mr/Ms "+surname)
    print("Ensure to pay " + client +" account the arrears or minimum installment.")
    print("Arrears: N$"+ arrears)
    print("Installment: N$" + installment)
    print("Reference: "+ref)
    print("Pay By "+ptp +"at nearest JDG store:Sleepmasters/Hi-Fi Corporation or Incredible Connection")
    print("Contact Resolve Account Collection Services:")
    print("tel: 061307491 0"+phone+"\n")
    print("-------------------------------------------------------------------------------------------------")
    
