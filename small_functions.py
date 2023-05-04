def get_excel_sheet_values_xlrd(sheet):
    sheet_values=[]
    for row in range(0,sheet.nrows):
        column=[]
        for col in range(0,sheet.ncols):
            column.append(sheet.cell(row,col).value)
        sheet_values.append(column)

    return sheet_values

def get_values_from_sheet_xlrd(workbook,sheetname):
    if sheetname not in workbook.sheet_names():
        print("Tab " + sheetname + " not found")
        return[0,[]]
    else:
        tab = workbook.sheet_by_name(sheetname)
        values = get_excel_sheet_values_xlrd(tab)
        return [1,values]