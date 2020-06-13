import MySQLdb as ms
from getpass import getpass

class DB:
    def __init__(self):
        # Get password for database but don't display on screen
        pswd = getpass("Password:")
        # Connect to database
        self.db = ms.connect(host="localhost", user="root", passwd=pswd, db="gregs_list")

    def RunQuery(self, query):
        # Select a query from a file
        # Execute a query to return from our database
        self.cur = self.db.cursor()
        self.cur.execute(query)

    def RunQueryFromFile(self, fileName):
        file = open(fileName, 'r')
        lines = file.readlines()
        query = ""
        for line in lines:
            query += line
        self.RunQuery(query)

    def PrintResult(self):
        ## Print out a table from our query
        print("")
        for row in self.cur.fetchall():
            print(row[0])

    def __del__(self):
        # Close out of our database
        self.db.close()
