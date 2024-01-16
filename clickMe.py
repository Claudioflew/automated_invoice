

import parseData
import makeInvoice
import openpyxl

MONTH_DICT = {
    1: "January", 2: "February", 3: "March", 4: "April",
    5: "May", 6: "June", 7: "July", 8: "August",
    9: "September", 10: "October", 11: "November", 12: "December"
    }

def main():

    print("\nParsing data... Takes up to 5 minutes.\n")

    startFromHere = "/Users/claudio/Desktop/CS_Projects/Automated_Invoice/StartFromHere.xlsx"
    id, month, year, invDate = getDateAndID(startFromHere)

    baseInfoDict, dataDict = parseData.parseData(id, month, year)



    print("Data parsing completed.\n")
    print("Making Invoices...\n")

    templatePath = "/Users/claudio/Desktop/CS_Projects/Automated_Invoice/invoiceTemplate.xlsx"
    outputPath = f"/Users/claudio/Desktop/{MONTH_DICT[month]} {year} Invoices.xlsx"
    
    makeInvoice.makeFile(templatePath, outputPath)
    makeInvoice.makeInvoice(outputPath, baseInfoDict, dataDict, invDate)

    li = outputPath.split("/")
    invoiceName = li[-1].rstrip(".xlsx")
    print(f"{invoiceName} was made!\n")
    

def getDateAndID(startFromHere):
    """
    getDateAndID loads an excel file and extracts an id of the sheet and date invoices will be issued on.

    :param startFromHere: string - absolute path of the excel file
    :return: valid id: int, month: int, year: int, invoicing date: string
    """
    wb = openpyxl.load_workbook(startFromHere)
    sheet = wb.active

    month = sheet["B2"].value
    date = sheet["B3"].value
    year = sheet["B4"].value
    id = sheet["B5"].value

    invDate = MONTH_DICT[month]
    invDate += f" {date}, {year}"

    month -= 1
    if month == 0:
        month = 12
        year -= 1

    return id, month, year, invDate

if __name__ == "__main__":
    main()