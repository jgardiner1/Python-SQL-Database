import mysql.connector

db = mysql.connector.connect(
    host="localhost",
    user="PerryDBTest",
    passwd="4MR&551hG2",
    database="Database1"
)

cursor = db.cursor(buffered=True)

# Creates new table
#cursor.execute("CREATE TABLE ContractorNew (name VARCHAR(50), service VARCHAR(50), email VARCHAR(50), contactNumber VARCHAR(20), responded BOOL, personID int PRIMARY KEY AUTO_INCREMENT)")

# describes table and creates iterable object you can loop through and print
#cursor.execute("DESCRIBE ContractorNew")

#for x in cursor:
#    print(x)

# inserts new element into table and commits change to the database. 
#cursor.execute("INSERT INTO ContractorNew (name, service, email, contactNumber, responded) VALUES (%s,%s,%s,%s,%s)", ("James Gardiner", "Programmer", "jamesgardiner1@live.co.uk", "07814844689", 1))
#db.commit()

#TABLE_NAME = "ContractorNew"

#cursor.execute(f"SELECT * FROM {TABLE_NAME}")

#for x in cursor:
#    print(x)

cursor.execute("DROP DATABASE Database1")
#cursor.execute("DROP TABLE Contractors")

db.commit()