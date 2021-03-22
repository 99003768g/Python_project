import pandas as pd
from openpyxl import load_workbook
from matplotlib import pyplot as pt

excel_file = pd.ExcelFile('book.xlsx')

something = input("ENTER PS or NAME or EMAIL as key: ")

if something == "PS":
    PS = int(input("Enter PS: "))
    df = pd.DataFrame()

    for i in excel_file.sheet_names:
        df1 = pd.read_excel(excel_file, i)
        # check for input
        df1.set_index('PS_Number', inplace=True)

        result = df1.loc[PS]
        print(result)
    excel_file.close()
    path = "out.xlsx"
    book = load_workbook(path)

    writer = pd.ExcelWriter(path, engine='openpyxl')
    writer.book = book
    if 'Mastersheet' in book.sheetnames:
        reference = book['Mastersheet']
        book.remove(reference)

    result.to_excel(writer, sheet_name='Mastersheet')
    pivot = df1.groupby(['PS_Number']).mean()
    variable = pivot.loc[:,"Training_Room_5":"Team_No_5"]
    variable.plot(kind ='bar')
    pt.show()
    writer.save()
    writer.close()

elif something == "NAME":
    NAME = input("Enter NAME: ")
    df = pd.DataFrame()

    for i in excel_file.sheet_names:
        df1 = pd.read_excel(excel_file, i)
        # check for input
        df1.set_index('Name', inplace=True)

        result = df1.loc[NAME]
        print(result)
    excel_file.close()
    path = "out.xlsx"
    book = load_workbook(path)

    writer = pd.ExcelWriter(path, engine='openpyxl')
    writer.book = book
    if 'Mastersheet' in book.sheetnames:
        reference = book['Mastersheet']
        book.remove(reference)

    result.to_excel(writer, sheet_name='Mastersheet')
    pivot = df1.groupby(['PS_Number']).mean()
    variable = pivot.loc[:,"Training_Room_5":"Team_No_5"]
    variable.plot(kind ='bar')
    pt.show()
    writer.save()
    writer.close()


elif something == "EMAIL":
    EMAIL = input("Enter Email: ")
    df = pd.DataFrame()

    for i in excel_file.sheet_names:
        df1 = pd.read_excel(excel_file, i)
        # check for input
        df1.set_index('Email_Address', inplace=True)

        result = df1.loc[EMAIL]
        print(result)
    excel_file.close()
    path = "out.xlsx"
    book = load_workbook(path)

    writer = pd.ExcelWriter(path, engine='openpyxl')
    writer.book = book
    if 'Mastersheet' in book.sheetnames:
        reference = book['Mastersheet']
        book.remove(reference)

    result.to_excel(writer, sheet_name='Mastersheet')
    pivot = df1.groupby(['PS_Number']).mean()
    variable = pivot.loc[:,"Training_Room_5":"Team_No_5"]
    variable.plot(kind ='bar')
    pt.show()
    writer.save()
    writer.close()

else:
    print("Valid keyword not entered Dude!")

excel_file.close()