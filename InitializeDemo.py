import ConnectToExcel as ctex
import ConnectToDB as ctdb

# Connect to database. "commit = True" means write to database.
db = ctdb.DB(commit = True)

# Drop Table
db.RunQuery("DROP TABLE IF EXISTS UserNames")
db.RunQuery("DROP TABLE IF EXISTS GroupTable")
db.RunQuery("DROP TABLE IF EXISTS GroupMembers")
                
# Get Excel files
GroupFile = ctex.ExcelFile("Groups.xlsx")
InputFile = ctex.ExcelFile("_Accounts.xlsx")

# Create table for Groups
f1 = GroupFile.CreateTableQuery(GroupFile.wb["GroupTable"])
f2 = GroupFile.CreateInsertQuery(GroupFile.wb["GroupTable"])
db.RunQueryFromFile("Queries/" + f1)
db.RunQueryFromFile("Queries/" + f2)
db.RunQuery('''ALTER TABLE GroupTable
               ADD PRIMARY KEY (GroupID);
            ''')

# Create table for Group table
f3 = InputFile.CreateTableQuery(InputFile.wb["UserNames"])
f4 = InputFile.CreateInsertQuery(InputFile.wb["UserNames"])
db.RunQueryFromFile("Queries/" + f3)
db.RunQueryFromFile("Queries/" + f4)
db.RunQuery('''ALTER TABLE UserNames
               ADD PRIMARY KEY (SecID);
            ''')

# Create GroupMembers Table
db.RunQuery('''CREATE TABLE IF NOT EXISTS GroupMembers
            (
                SecID INT,
                GroupID INT,
                GroupMemberID INT PRIMARY KEY AUTO_INCREMENT
             );
            ''')
