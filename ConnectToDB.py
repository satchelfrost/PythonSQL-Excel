import MySQLdb as ms
import json
from tabulate import tabulate
from getpass import getpass

class DB:
    def __init__(self, commit=True):


        # Get database settings from file
        with open("DB_SETTINGS.txt") as json_file:
            dbSettings = json.load(json_file)

        pswd = ''
        if dbSettings['pswd'] == '':
            # Get password for database but don't display on screen
            pswd = getpass("Password:")
        else:
            pswd = dbSettings['pswd']
        
        
        # Connect to database
        self.db = ms.connect(host=dbSettings['host'],
                             user=dbSettings['user'],
                             passwd=pswd,
                             db=dbSettings['db'])
        self.commit = commit

    def RunQuery(self, query):
        # Select a query from a file
        # Execute a query to return from our database
        self.cur = self.db.cursor()
        self.cur.execute(query)
        if self.commit:
            self.db.commit()

    def RunQueryFromFile(self, fileName):
        file = open(fileName, 'r')
        lines = file.readlines()
        query = ""
        for line in lines:
            query += line
        if query != "blank":
            self.RunQuery(query)

    def PrintResult(self):
        ## Print out a table from our query
        Headers = []
        if self.cur.description is not None:
            for des in self.cur.description:
                Headers.append(des[0])
        print(tabulate(self.cur.fetchall(), headers=Headers, tablefmt='orgtbl'))

    def __del__(self):
        # Close out of our database
        self.db.close()

