import ConnectToDB as ctdb
import ConnectToExcel as ctex

# Get an instance of our database
db = ctdb.DB()

# Get an instance of our spreadsheet
ss = ctex.SpreadSheet("example.xlsx")

# Generate queries and save in files
for sheet in ss.wb.worksheets:
    ss.GenerateInsertQuery(sheet, sheet.title)
    ss.CreateTableQuery(sheet, sheet.title)
