# Automation for python and SQL

### Introduction
This project aims to read in tables from an Excel file and generate queries for MySQL. The assumption is made that the excel file is layed out in the same way as the example.xlsx given. The assumption is also made that you know how to connect to a MySQL database (i.e. information such as host, user, pswd, database, is known to you).

### Install necessary libraries
To start off, first make sure the proper libraries are installed for Python.
First make sure the latest version of pip is installed
```shell
python -m pip install --upgrade pip
```

Install library for interacting with SQL
```shell
pip install mysqlclient
```
Install library for interacting with Microsoft Excel
```shell
pip install openpyxl
```

### Cloning the repository
Once you have installed all necessary libraries you can clone the repository either by clicking clone on github, or directly in the terminal assuming you have git installed and it's in your PATH:

```shell
git clone https://github.com/satchelfrost/PythonSQL-Excel
```

### Database Settings
Next you're going to want to change the DB_Settings.txt to whatever your particular settings. For host, "localhost" will work if you're just testing this out on your local machine. For pswd you may type in your password or simply leave this field blank (e.g. "pswd" : "") and you will be prompted for a password if you run the demo.py script from the terminal (the password input will be hidden). Note however that if you run the script from Python's Idle, the password will be displayed as you type it in. Finally, it should be obvious, but for the "db" field, if the database does not already exist you will not be able to make a connection.
