import ConnectToExcel as ctex
import ConnectToDB as ctdb

# Load file with group information from different sites
GroupFile = ctex.ExcelFile("Groups.xlsx")

# Load input file
InputFile = ctex.ExcelFile("_Accounts.xlsx")

# Connect to database. "commit = True" means write to database.
db = ctdb.DB(commit = True)

# Get the worksheet of user account info
ws = InputFile.wb["UserInfo"]

# Generate the groupmember insertion query and run
GroupFile.CreateGroupMemberInsertQuery(ws, db)
GroupFile.EliminateDuplicates("Queries/Insertion.txt")
db.RunQueryFromFile("Queries/Insertion.txt")

# Generate the deletion query if any and run
GroupFile.CreateGroupMemberDeletionQuery(ws, db)
db.RunQueryFromFile("Queries/GroupMemberDeletion.txt")

# Deletion query for deleting username
GroupFile.CreateUsernameDeletionQuery(ws)
db.RunQueryFromFile("Queries/UsernameDeletion.txt")

