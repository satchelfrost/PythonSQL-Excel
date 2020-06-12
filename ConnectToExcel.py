import openpyxl as opx
from tabulate import tabulate

# Load the excel file, and the specific sheet
wb = opx.load_workbook("example.xlsx")

def CreateListFromWorksheet(ws):
    MyList = []
    for i in range(ws.max_row-1):
        MyList.append([])

    for i in range(2, ws.max_row + 1):
        for j in range(1, ws.max_column + 1):
            CellValue = ws.cell(row = i, column = j).value
            MyList[i-2].append(CellValue)
    return MyList

def CreateHeadersFromWorksheet(ws):
    Headers = []
    for i in range(1, ws.max_column + 1):
        CellValue = ws.cell(row = 1, column = i).value
        Headers.append(CellValue)
    return Headers

def PrintTableFromWorksheet(ws):
    TableData = CreateListFromWorksheet(ws)
    Headers = CreateHeadersFromWorksheet(ws)
    print(tabulate(TableData, headers=Headers, tablefmt='orgtbl'))

for sheet in wb.worksheets:
    PrintTableFromWorksheet(sheet)
    print("")
    


    


