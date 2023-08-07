import logging
import json
import Database
import Application
import GUI

## TODO
# implement deselect all
# reset scrollbar when navigating pages or performing another query
# change string concatentation of SQL queries to prevent injection attacks

## MAIN CODE
log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
logger = logging.getLogger(__name__)
logger.setLevel('DEBUG')
file_handler = logging.FileHandler('Logs.log')
formatter = logging.Formatter(log_format)
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)

# Reading user configuration
with open('information.txt') as f:
    data = f.read()
    js = json.loads(data)
    f.close()

# Constants from json file
HOST = js["HOST"]
USER = js["USER"]
PASSWD = js["PASSWD"]
DATABASE = js["DATABASE"]
TABLE_NAME = js["TABLE_NAME"]
APP_NAME = js["APP_NAME"]
OUTLOOK_LOC = js["OUTLOOK_LOC"]
MAX_RESULTS_PPAGE = js["MAX_RESULTS_PPAGE"]
TEST_DATA = js["TEST_DATA"]
DEL_DATABASE = js["DEL_DATABASE"]
DEL_TABLE = js["DEL_TABLE"]

db = None
while (db == None):
    db = Database.connect_database(host=HOST, user=USER, passwd=PASSWD, database=DATABASE, table_name=TABLE_NAME)


cursor = db.cursor(buffered=True)

if DEL_TABLE == "True":
    Database.delete_table(cursor=cursor, table=TABLE_NAME)
if DEL_DATABASE == "True":
    Database.delete_database(cursor=cursor, db=db, database=DATABASE)
    print("Table and Database sucessfully deleted. Exiting")
    exit(0)

if js["READ_TEST_DATA"] == "True":
    Database.read_test_data(db=db, cursor=cursor, table_name=TABLE_NAME, test_data=TEST_DATA)

if js["OPEN_OUTLOOK"] == "True":
    Application.open_outlook(outlook_loc=OUTLOOK_LOC)

app = GUI.App(APP_NAME=APP_NAME, TABLE_NAME=TABLE_NAME, MAX_RESULTS_PPAGE=MAX_RESULTS_PPAGE, cursor=cursor, db=db)
app.mainloop()

Application.close_outlook()