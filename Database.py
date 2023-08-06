import mysql.connector
import logging
import csv

log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
logger = logging.getLogger(__name__)
logger.setLevel('DEBUG')
file_handler = logging.FileHandler('Logs.log')
formatter = logging.Formatter(log_format)
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)


def connect_database(host: str, user: str, passwd: str, database: str, table_name: str) -> mysql.connector.connection:
    # Database Connection
    try:
        db = mysql.connector.connect(
            host=host,
            user=user,
            passwd=passwd,
            database=database
        )
        logger.info('{}'.format(f"Successfully connected to database: {database}"))
        return db
    except mysql.connector.Error as e:
            logger.error('{}'.format(f"ERROR: {e.errno} - SQLSTATE value: {e.sqlstate} - Error Message: {e.msg}"))
            logger.error('{}'.format(f"Attempting to create database {database}"))
            return create_database(host=host, user=user, passwd=passwd, database=database, table_name=table_name)


def create_database(host: str, user: str, passwd: str, database: str, table_name: str) -> mysql.connector.connection:
    try:
        db = mysql.connector.connect(
            host=host,
            user=user,
            passwd=passwd
        )

        cursor = db.cursor(buffered=True)

        cursor.execute(f"CREATE DATABASE {database}")
        logger.info('{}'.format(f"Successfully created database: {database}"))

        db = mysql.connector.connect(
                host=host,
                user=user,
                passwd=passwd,
                database=database
            )
        cursor = db.cursor(buffered=True)

        cursor.execute(f"CREATE TABLE {table_name} (name VARCHAR(50), service VARCHAR(50), email VARCHAR(50), contactNumber VARCHAR(30), responded BOOL, personID int PRIMARY KEY AUTO_INCREMENT)")
        logger.info('{}'.format(f"Successfully created table: {table_name}"))

        return db
    except mysql.connector.Error as e:
        logger.error('{}'.format(f"ERROR: {e.errno} - SQLSTATE value: {e.sqlstate} - Error Message: {e.msg}"))


def read_test_data(db, cursor, table_name, test_data):
    logger.info('{}'.format(f"Attempting to read {test_data}"))
    try:
        file = open(test_data, 'r')
        reader = csv.reader(file)
        counter = 0

        for record in reader:
            cursor.execute(f"INSERT INTO {table_name} (name, service, email, contactNumber, responded) VALUES (%s,%s,%s,%s,%s)", (f"{record[0]}", f"{record[1]}", f"{record[2]}", f"{record[3]}", f"0"))
            db.commit()
            counter += 1
        
        file.close()

        logger.info('{}'.format(f"Successfully read {counter} files into {table_name}"))
    except FileNotFoundError as e:
        logger.error('{}'.format(f"ERROR: {e.errno} - MESSAGE: {e.strerror}"))
    except mysql.connector.Error as e:
        logger.error('{}'.format(f"ERROR: {e.errno} - SQLSTATE value: {e.sqlstate} - Error Message: {e.msg}"))


def delete_database(cursor, db, database):
    try:
        cursor.execute(f"DROP DATABASE {database}")
        db.commit()
        logger.info('{}'.format(f"Successfully deleted database: {database}"))
    except mysql.connector.Error as e:
        logger.error('{}'.format(f"ERROR: {e.errno} - SQLSTATE value: {e.sqlstate} - Error Message: {e.msg}"))
        logger.error('{}'.format(f"Could not delete database: {database}"))


def delete_table(cursor, table):
    try:
        cursor.execute("DROP TABLE {table}}")
        logger.info('{}'.format(f"Successfully deleted table: {table}"))
    except mysql.connector.Error as e:
            logger.error('{}'.format(f"ERROR: {e.errno} - SQLSTATE value: {e.sqlstate} - Error Message: {e.msg}"))
            logger.error('{}'.format(f"Could not delete table: {table}"))
