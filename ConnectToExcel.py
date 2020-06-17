import openpyxl as opx
import os
from tabulate import tabulate

class ExcelFile:
    
    def __init__(self, name):
        # Load the excel file, and the specific sheet
        self.wb = opx.load_workbook(name)

        # make a Queries folder if it doesn't exist
        if not os.path.exists('Queries'):
            os.mkdir("Queries")

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

    def CreateTableQuery(self, ws):
        # Insert portion of query
        file = open("Queries/" + ws.title + "_TableCreation.txt", 'w')
        string = "CREATE TABLE IF NOT EXISTS " + ws.title
        file.write(string)
        string = ""

        # Header field portion of query        
        Headers = self.CreateHeadersFromWorksheet(ws)
        fields = "\n(\n"
        for i in range(len(Headers)):
            fields += "\t" + Headers[i]
            Cell = ws.cell(row = 2, column = i + 1).value
            if isinstance(Cell, str):
                fields += " VARCHAR(30)"
            if isinstance(Cell, int):
                fields += " INT"
            if isinstance(Cell, float):
                fields += " DEC(4,2)" 
            if (i != len(Headers) - 1):
                fields += ", "
            fields += "\n"
        string += fields + ");"
        file.write(string)
        string = ""
        return ws.title + "_TableCreation.txt"

    def CreateInsertQuery(self, ws):
        # Insert portion of query
        file = open("Queries/" + ws.title + "_Insertion.txt", 'w')
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
                if isinstance(TableData[i][j], str):
                    value += "'"
                value += str(TableData[i][j])
                if isinstance(TableData[i][j], str):
                    value += "'"
                if (j != len(TableData[i]) - 1):
                    value += ", "
            value += ")"
            if (i != len(TableData) - 1):
                    value += ",\n"
            values += value
        string += values
        string += ";"
        file.write(string)
        return ws.title + "_Insertion.txt"
