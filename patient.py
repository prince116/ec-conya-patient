# Description: mini program for monthly task in my company

from datetime import datetime as dt
from pandas.io.formats.excel import ExcelFormatter
import os
import pandas as pd

ExcelFormatter.header_style = None
outputFilename: str = "output_" + dt.now().strftime("%Y%m%d_%H%M%S") + ".xlsx"


def processFile(filepath: str):
    df = pd.read_excel(filepath, header=None)
    pd_data = df[1:len(df)]

    dataList: list = []
    header: list = [
        "Patient No.",
        "Surname",
        "Given Name",
        "Surname (Chinese)",
        "Given Name (Chinese)",
        "Sex",
        "DOB (YYYY-MM-DD)",
        "HKID/Passport No.",
        "Mobile",
        "Home",
        "Office",
        "Fax",
        "Email",
        "Address",
        "Preferred/Mailing Address"
    ]

    for row, col in pd_data.iterrows():

        patientNo: str = col[0]
        surname: str = col[1]
        givenName: str = col[2]
        surnameChinese: str = col[3]
        givenNameChinese: str = col[4]
        sex: str = col[5]

        try:
            dob: str = dt.strptime(col[6], "%d/%m/%Y").strftime("%Y-%m-%d")
        except ValueError:
            dob: str = "wrong_format"

        identity: str = col[7]
        mobile: str = col[8]
        homeTel: str = col[9]
        officeTel: str = col[10]
        fax: str = col[11]
        email: str = col[12]
        address: str = col[13]
        mailingAddress: str = col[14]

        patientData: list = [patientNo, surname, givenName, surnameChinese, givenNameChinese, sex, dob, identity, mobile, homeTel, officeTel, fax, email, address, mailingAddress]
        dataList.append(patientData)

        print("[%s/%s]" % (row, len(pd_data)), end='\r')

    excelDf = pd.DataFrame(dataList)
    with pd.ExcelWriter(outputFilename) as writer:
        print("Writing data to the Excel file ...")
        excelDf.to_excel(writer, sheet_name="Sheet1", index=False, header=header)

        workbook = writer.book
        worksheet = writer.sheets["Sheet1"]

        # Set cell format
        cellFormat = workbook.add_format()
        cellFormat.set_align('left')
        worksheet.set_column('A:O', None, cellFormat)

        # Set header format
        headerFormat = workbook.add_format()
        headerFormat.set_bold()
        headerFormat.set_italic()
        headerFormat.set_align("left")
        headerFormat.set_fg_color("#99CCFF")
        headerFormat.set_font_size(10)
        headerFormat.set_font_name("Tahoma")

        for index, col in enumerate(header):
            worksheet.write(0, index, col, headerFormat)

        writer.save()
        print("Finished!")


def askForFileName():
    fileName: str = input("Enter File Name: ")
    filePath: str = os.path.join(os.path.abspath(os.getcwd()), fileName)

    if not os.path.exists(filePath):
        print("File dost not exist.")

    processFile(filePath)


def main():
    askForFileName()


if __name__ == "__main__":
    main()
