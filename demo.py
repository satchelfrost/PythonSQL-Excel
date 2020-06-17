import ConnectToExcel as ctex
import ConnectToDB as ctdb

# Load the excel file
file = ctex.ExcelFile("example.xlsx")

# Connect to database
db = ctdb.DB(commit = True)

# Make and run queries
for sheet in file.wb.worksheets:
    TableQueryFileName = file.CreateTableQuery(sheet)
    InsertQueryFileName = file.CreateInsertQuery(sheet)
    db.RunQueryFromFile("Queries/" + TableQueryFileName)
    db.RunQueryFromFile("Queries/" + InsertQueryFileName)
