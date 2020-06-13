import openpyxl as opx
from tabulate import tabulate

class SpreadSheet:
    
    def __init__(self, name):
        # Load the excel file, and the specific sheet
        self.wb = opx.load_workbook(name)

    # This list will hold the data from the spreadsheet
    def CreateTableDataFromWorksheet(self, ws):
        TableData = []
        for i in range(ws.max_row-1):
            TableData.append([])
            
        for i in range(2, ws.max_row + 1):
            for j in range(1, ws.max_column + 1):
                CellValue = ws.cell(row = i, column = j).value
                TableData[i-2].append(CellValue)
        return TableData

    # Headers from th spreadsheet
    def CreateHeadersFromWorksheet(self, ws):
        Headers = []
        for i in range(1, ws.max_column + 1):
            CellValue = ws.cell(row = 1, column = i).value
            Headers.append(CellValue)
        return Headers

    # Print out the table from the worksheet
    def PrintTableFromWorksheet(self, ws):
        TableData = self.CreateTableDataFromWorksheet(ws)
        Headers = self.CreateHeadersFromWorksheet(ws)
        print(tabulate(TableData, headers=Headers, tablefmt='orgtbl'))

    def PrintAllTables(self):
        for sheet in self.wb.worksheets:
            self.PrintTableFromWorksheet(sheet)
            print("")

    def GenerateInsertQuery(self, ws, name):
        # Insert portion of query
        file = open("Queries/" + name + ".txt", 'w')
        string = "INSERT INTO " + ws.title
        file.write(string)
        string = ""

        # Header field portion of query        
        Headers = self.CreateHeadersFromWorksheet(ws)
        fields = "\n("
        for i in range(len(Headers)):
            fields += Headers[i]
            if (i != len(Headers) - 1):
                fields += ", "
        string += fields + ")"
        file.write(string)
        string = ""

        # Beginning of values portion of query
        file.write("\nVALUES\n")

        # Values portion of query
        TableData = self.CreateTableDataFromWorksheet(ws)
        values = ""
        for i in range(len(TableData)):
            value = "("
            for j in range(len(TableData[i])):
                value += str(TableData[i][j])
                if (j != len(TableData[i]) - 1):
                    value += ", "
            value += ")"
            if (i != len(TableData) - 1):
                    value += ",\n"
            values += value
        string += values
        file.write(string)
        
ss = SpreadSheet("example.xlsx")
for sheet in ss.wb.worksheets:
    ss.GenerateInsertQuery(sheet, sheet.title)

