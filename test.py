import win32com.client
import mysql.connector
import logging

logging.basicConfig(filename='Logs.log', level=logging.INFO, format='%(asctime)s:%(message)s')

def create_database():
    db = mysql.connector.connect(
        host="localhost",
        user="PerryDBTest",
        passwd="4MR&551hG2")
    
    cursor = db.cursor()

    cursor.execute("CREATE DATABASE testdatabase2")

def create_table():
     cursor.execute("CREATE TABLE ContractorNew2 (name VARCHAR(50), service VARCHAR(50), email VARCHAR(50), contactNumber VARCHAR(20), responded BOOL, personID int PRIMARY KEY AUTO_INCREMENT)")


# Database Connection
try:
    db = mysql.connector.connect(
        host="localhost",
        user="PerryDBTest",
        passwd="4MR&551hG2",
        database="testdatabase2"
    )
except mysql.connector.Error as e:
        print("Error code: ", e.errno,
              "\nSQLSTATE value: ", e.sqlstate,
              "\nError Message: ", e.msg)

        create_database()
        
cursor = db.cursor(buffered=True)

try:
    cursor.execute("INSERT INTO ContractorNew2 (name, service, email, contactNumber, responded) VALUES (%s,%s,%s,%s,%s)", ("James Gardiner", "Programmer", "jamesgardiner1@live.co.uk", "07814844689", 1))
except mysql.connector.errors.ProgrammingError as e:
    print("Error code: ", e.errno,
        "\nSQLSTATE value: ", e.sqlstate,
        "\nError Message: ", e.msg)
    
    create_table()