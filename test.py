import MySQLdb as ms
from getpass import getpass

# Get password for database but don't display on screen
pswd = getpass("Password:")

# Connect to database
db = ms.connect(host="localhost", user="root", passwd=pswd, db="gregs_list")

# Select a query from a file
q = open('Query_00.txt', 'r')
lines = q.readlines()
query = ""
for line in lines:
	query += line 

# Execute a query to return from our database
cur = db.cursor()
cur.execute(query)

## Print out a table from our query
print("")
for row in cur.fetchall():
        print(row[0])

# Close out of our database
db.close()

