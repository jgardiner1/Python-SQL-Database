import customtkinter as ctk
from tkinter import messagebox
import win32com.client
import mysql.connector
from functools import partial
import os
import logging
import csv
import json
import time
from PIL import Image

## TODO
# Finish implementing checkbox deletion of results to streamline process
# implementing by repeating search query after attempting to delete and recollecting results instead of ammending current

class App(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("1280x720")
        self.title(APP_NAME)
        self.GLOBAL_RESULTS = []
        self.REMOVAL_LIST = []
        self.MAX_PAGES = 0
        self.CURRENT_PAGE = 1
        self.RESULTS_PER_PAGE = MAX_RESULTS_PPAGE
        self.LAST_QUERY = ""

        # Setting window appearances
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        # Clears result frame when update is needed
        def clear_frame():
            print("\nClearing frame")
            start = time.perf_counter()
            for widget in frameResults.winfo_children():
                widget.destroy()

            end=time.perf_counter()
            print("Time to clear frame: ", end-start)
        

        def repeat_search():
            cursor.execute(self.LAST_QUERY)
            results = cursor.fetchall()

            # Calculate max pages and configure page 1
            if len(results) % self.RESULTS_PER_PAGE != 0:
                self.MAX_PAGES = (len(results) // self.RESULTS_PER_PAGE) + 1
            else:
                self.MAX_PAGES = (len(results) // self.RESULTS_PER_PAGE)
            currentPage.configure(text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
            
            # Split results into array of lists for each page
            self.GLOBAL_RESULTS = [results[x:x+self.RESULTS_PER_PAGE] for x in range(0, len(results), self.RESULTS_PER_PAGE)]
            
            clear_frame()
            load_results(self.CURRENT_PAGE - 1)
            return

        # Logic for adding new data into the database
        def button_event_add():
            # ensures service is selected when inputting new entry
            if serviceEntry.get() == "Select Service":
                messagebox.showerror('ERROR', 'Please make sure the service type field is filled in.')
                return
            
            # try/catch for entering new data
            try:
                cursor.execute(f"INSERT INTO {TABLE_NAME} (name, service, email, contactNumber, responded) VALUES (%s,%s,%s,%s,%s)", (f"{nameEntry.get()}", f"{serviceEntry.get()}", f"{emailEntry.get()}", f"{contactEntry.get()}", f"{respondedCheckBox.get()}"))
                db.commit()
                logging.debug('{}'.format(f"Successfully insert into: {TABLE_NAME} - INFO: {nameEntry.get()}, {serviceEntry.get()}, {emailEntry.get()}, {contactEntry.get()}, {respondedCheckBox.get()}"))
            except mysql.connector.Error as e:
                logging.debug('{}'.format(f"ERROR: {e.errno} - SQLSTATE value: {e.sqlstate} - Error Message: {e.msg}"))
            
            # resetting all fields
            nameEntry.delete(0, 50)
            serviceEntry.set("None")
            emailEntry.delete(0, 50)
            contactEntry.delete(0, 50)
            respondedCheckBox.deselect()


        def button_event_search():
            self.CURRENT_PAGE = 1
            # base query. Selects everything
            query = f"SELECT * FROM {TABLE_NAME}"
            conditions = []

            service = serviceSearch.get()
            name = nameSearch.get()

            # If checks to construct SQL query to execute
            if service != "None" and service != "Search by Service":
                conditions.append(f"service='{service}'")

            if name != "":
                conditions.append(f"name LIKE '%{name}%'")

            if respondedSearch.get() == 1:
                conditions.append(f"responded=1")
            
            if len(conditions) == 1:
                query += f" WHERE {conditions[0]}"
            if len(conditions) > 1:
                query += f" WHERE {conditions[0]}"
                for c in range(1, len(conditions)):
                    query += f" AND {conditions[c]}"
            
            if alphabeticalCheck.get() == 1:
                query += f" ORDER BY name"
            
            # Execute query and store results
            cursor.execute(query)
            self.LAST_QUERY = query
            results = cursor.fetchall()

            # Incase new query yields no results
            if len(results) == 0:
                clear_frame()
                ctk.CTkLabel(master=frameResults, text="No Results...").pack()
                return
            
            # Calculate max pages and configure page 1
            if len(results) % self.RESULTS_PER_PAGE != 0:
                self.MAX_PAGES = (len(results) // self.RESULTS_PER_PAGE) + 1
            else:
                self.MAX_PAGES = (len(results) // self.RESULTS_PER_PAGE)
            currentPage.configure(text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
            
            # Split results into array of lists for each page
            self.GLOBAL_RESULTS = [results[x:x+self.RESULTS_PER_PAGE] for x in range(0, len(results), self.RESULTS_PER_PAGE)]
            
            clear_frame()
            load_results(0)
        

        def button_event_delete():
            removal_list = ','.join(str(int(x)) for x in self.REMOVAL_LIST)

            # Selects individual from database, deletes and logs
            try:
                cursor.execute(f"DELETE FROM {TABLE_NAME} WHERE personID IN ({removal_list})")
                db.commit()
                logging.debug('{}'.format(f"Successfully deleted {cursor.rowcount} entries from: {TABLE_NAME}"))
                self.REMOVAL_LIST.clear()
                repeat_search()
            except mysql.connector.Error as e:
                logging.debug('{}'.format(f"ERROR: {e.errno} - SQLSTATE value: {e.sqlstate} - Error Message: {e.msg}"))
            return


        def checkbox_event_entry_selection(id):
            if id in self.REMOVAL_LIST:
                self.REMOVAL_LIST.remove(id)
                print("removed")
            else:
                self.REMOVAL_LIST.append(id)
                print("added")


        def button_event_email_open(emailAddress):
            outlook = win32com.client.Dispatch('Outlook.Application')
            email = outlook.CreateItem(0)
            email.To = emailAddress
            email.Display(True)

        
        def button_event_add_service():
            service = addServiceEntry.get()
            services = button_event_reload_services()
            if service in services:
                messagebox.showerror('ERROR', 'Service already within list')
                addServiceEntry.delete(0, 50)
            else:
                file = open('services.txt', 'a')
                file.write(f"\n{service}")
                file.close()

                services = button_event_reload_services()
                addServiceEntry.delete(0, 50)


        def button_event_reload_services():
            # Reading available services
            with open('services.txt') as f:
                services = [l for l in (line.strip() for line in f) if l]
                services.insert(0, "None")
                f.close()
            
            serviceEntry.configure(values=services)
            serviceSearch.configure(values=services)
            return services
        
        
        def button_event_edit_services():
            os.system('services.txt')
            return
        
        
        def button_event_page_down():
            if self.CURRENT_PAGE > 1:
                self.CURRENT_PAGE -= 1
                currentPage.configure(text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
                clear_frame()
                load_results(self.CURRENT_PAGE - 1)
            return
        

        def button_event_page_up():
            if self.CURRENT_PAGE < self.MAX_PAGES:
                self.CURRENT_PAGE += 1
                currentPage.configure(text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
                clear_frame()
                load_results(self.CURRENT_PAGE - 1)
            return


        def load_results(pageNum):
            print("Loading. Starting counter")
            start = time.perf_counter()
            for x in range(len(self.GLOBAL_RESULTS[pageNum])):
                resultFrame = ctk.CTkFrame(master=frameResults)
                resultFrame.pack(padx=5, pady=3, fill=ctk.BOTH, expand=True)

                # Results
                if self.GLOBAL_RESULTS[pageNum][x][5] in self.REMOVAL_LIST:
                    temp = ctk.CTkCheckBox(master=resultFrame, text=None, width=0, command=partial(checkbox_event_entry_selection, self.GLOBAL_RESULTS[pageNum][x][5]))
                    temp.grid(row=x, column=1, padx=10, pady=5)
                    temp.select()
                else:
                    ctk.CTkCheckBox(master=resultFrame, text=None, width=0, command=partial(checkbox_event_entry_selection, self.GLOBAL_RESULTS[pageNum][x][5])).grid(row=x, column=1, padx=10, pady=5)
                ctk.CTkLabel(master=resultFrame, corner_radius=0, text=self.GLOBAL_RESULTS[pageNum][x][0], width=150).grid(row=x, column=2, padx=10, pady=5)
                ctk.CTkLabel(master=resultFrame, corner_radius=0, text=self.GLOBAL_RESULTS[pageNum][x][1], width=150).grid(row=x, column=3, padx=10, pady=5)
                ctk.CTkLabel(master=resultFrame, corner_radius=0, text=self.GLOBAL_RESULTS[pageNum][x][2], width=200).grid(row=x, column=4, padx=10, pady=5)
                ctk.CTkLabel(master=resultFrame, corner_radius=0, text=self.GLOBAL_RESULTS[pageNum][x][3], width=150).grid(row=x, column=5, padx=10, pady=5)

                # Delete and Open Email Buttons
                #ctk.CTkButton(master=resultFrame, text="Delete", width=70, command=partial(button_event_delete, self.GLOBAL_RESULTS[pageNum][x][5], x)).grid(row=x, column=5, padx=2, pady=5, sticky=ctk.E)
                ctk.CTkButton(master=resultFrame, text="Open Email", width=80, command=partial(button_event_email_open, self.GLOBAL_RESULTS[pageNum][x][2])).grid(row=x, column=6, padx=2, pady=5, sticky=ctk.E)
            end = time.perf_counter()
            print("Time to load Results: ", end-start)


        ## RESULTS FRAME
        frameRight = ctk.CTkFrame(master=self)
        frameRight.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True, side=ctk.RIGHT)

        # Title
        ctk.CTkLabel(master=frameRight, text="RESULTS", fg_color="transparent", font=("Barlow Condensed", 25)).pack(pady=7)

        # Frame holds results
        frameResults = ctk.CTkScrollableFrame(master=frameRight)
        frameResults.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True)

        print(frameResults.winfo_height())

        # Frame for under main results frame. Stores page selection and results per page
        frameBottomRight = ctk.CTkFrame(master=frameRight, fg_color="gray13")
        frameBottomRight.pack(padx=5, pady=5, fill=ctk.X, expand=False)

        ctk.CTkButton(master=frameBottomRight, text="DELETE SELECTED RESULTS", command=button_event_delete).pack(padx=10, pady=10, side=ctk.LEFT)

        # Page Selection frame to store buttons and current page info
        framePageSelection = ctk.CTkFrame(master=frameBottomRight)
        framePageSelection.pack(padx=10, pady=10, fill=None, expand=True, side=ctk.LEFT, anchor=ctk.CENTER)
        ctk.CTkButton(master=framePageSelection, text="<", command=button_event_page_down).pack(padx=10, pady=10, side=ctk.LEFT)
        currentPage = ctk.CTkLabel(master=framePageSelection, text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
        currentPage.pack(padx=10, pady=10, side=ctk.LEFT)
        ctk.CTkButton(master=framePageSelection, text=">", command=button_event_page_up).pack(padx=10, pady=10, side=ctk.LEFT)


        ## SEARCH FRAME
        frameMiddleLeft = ctk.CTkFrame(master=self)
        frameMiddleLeft.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True)

        # Title
        ctk.CTkLabel(master=frameMiddleLeft, text="SEARCH", fg_color="transparent", font=("Barlow Condensed", 25)).pack(pady=7)

        serviceSearch = ctk.CTkOptionMenu(master=frameMiddleLeft, values=services, width=200)
        serviceSearch.set("Search by Service")
        serviceSearch.pack(pady=5, padx=20)

        nameSearch = ctk.CTkEntry(master=frameMiddleLeft, placeholder_text="Search by Name", width=200)
        nameSearch.pack(pady=5, padx=20)

        respondedSearch = ctk.CTkCheckBox(master=frameMiddleLeft, text="Filter by Responded")
        respondedSearch.pack(pady=5, padx=10)

        alphabeticalCheck = ctk.CTkCheckBox(master=frameMiddleLeft, text="Order Alphabetically")
        alphabeticalCheck.pack(padx=5, pady=10)

        # Search Button
        ctk.CTkButton(master=frameMiddleLeft, text="SEARCH", command=button_event_search).pack(pady=10, padx=20, side=ctk.BOTTOM)


        ## NEW ENTRIES FRAME
        frameTopLeft = ctk.CTkFrame(master=self)
        frameTopLeft.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True)

        # Title
        ctk.CTkLabel(master=frameTopLeft, text="NEW ENTRIES", fg_color="transparent", font=("Barlow Condensed", 25)).pack(pady=7)

        nameEntry = ctk.CTkEntry(master=frameTopLeft, placeholder_text="Name", width=200)
        nameEntry.pack(pady=5, padx=20)

        serviceEntry = ctk.CTkOptionMenu(master=frameTopLeft, values=services, width=200)
        serviceEntry.set("Select Service")
        serviceEntry.pack(pady=5, padx=20)

        emailEntry = ctk.CTkEntry(master=frameTopLeft, placeholder_text="Email", width=200)
        emailEntry.pack(pady=5, padx=20)

        contactEntry = ctk.CTkEntry(master=frameTopLeft, placeholder_text="Contact Number", width=200)
        contactEntry.pack(pady=5, padx=20)

        respondedCheckBox = ctk.CTkCheckBox(master=frameTopLeft, text="Responded?")
        respondedCheckBox.pack(pady=5, padx=20)

        # Add Button
        ctk.CTkButton(master=frameTopLeft, text="ADD", command=button_event_add).pack(side="bottom", pady=10, padx=20)


        ## SERVICES FRAME
        frameBottomLeft = ctk.CTkFrame(master=self)
        frameBottomLeft.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True)

        # Title
        ctk.CTkLabel(master=frameBottomLeft, text="EDIT SERVICES", fg_color="transparent", font=("Barlow Condensed", 25)).pack(pady=7)

        addServiceFrame = ctk.CTkFrame(master=frameBottomLeft, fg_color="gray13")
        addServiceFrame.pack(padx=5, pady=5, fill=ctk.BOTH, expand=True)
        addServiceEntry = ctk.CTkEntry(master=addServiceFrame, placeholder_text="New Service")
        addServiceEntry.pack(padx=5, pady=5, side=ctk.LEFT)
        ctk.CTkButton(master=addServiceFrame, text="ADD", command=button_event_add_service).pack(padx=5, pady=5, side=ctk.LEFT)
        
        # Reload and Edit services button
        editServiceFrame = ctk.CTkFrame(master=frameBottomLeft, fg_color="gray13")
        editServiceFrame.pack(padx=5, pady=5, fill=ctk.Y, expand=True)
        ctk.CTkButton(master=editServiceFrame, text="RELOAD SERVICES", command=button_event_reload_services).pack(padx=5, pady=5, side=ctk.BOTTOM)
        ctk.CTkButton(master=editServiceFrame, text="EDIT SERVICES", command=button_event_edit_services).pack(padx=5, pady=5, side=ctk.BOTTOM)


def connect_database():
    # Database Connection
    try:
        db = mysql.connector.connect(
            host=HOST,
            user=USER,
            passwd=PASSWD,
            database=DATABASE
        )
        logging.debug('{}'.format(f"Successfully connected to database: {DATABASE}"))
        return db
    except mysql.connector.Error as e:
            logging.debug('{}'.format(f"ERROR: {e.errno} - SQLSTATE value: {e.sqlstate} - Error Message: {e.msg}"))
            logging.debug('{}'.format(f"Attempting to create database {DATABASE}"))
            return create_database()


def create_database():
    try:
        db = mysql.connector.connect(
            host=HOST,
            user=USER,
            passwd=PASSWD
        )

        cursor = db.cursor(buffered=True)

        cursor.execute(f"CREATE DATABASE {DATABASE}")
        logging.debug('{}'.format(f"Successfully created database: {DATABASE}"))

        db = mysql.connector.connect(
                host=HOST,
                user=USER,
                passwd=PASSWD,
                database=DATABASE
            )
        cursor = db.cursor(buffered=True)

        cursor.execute(f"CREATE TABLE {TABLE_NAME} (name VARCHAR(50), service VARCHAR(50), email VARCHAR(50), contactNumber VARCHAR(30), responded BOOL, personID int PRIMARY KEY AUTO_INCREMENT)")
        logging.debug('{}'.format(f"Successfully created table: {TABLE_NAME}"))

        return db
    except mysql.connector.Error as e:
        logging.debug('{}'.format(f"ERROR: {e.errno} - SQLSTATE value: {e.sqlstate} - Error Message: {e.msg}"))


def read_test_data():
    logging.debug('{}'.format(f"Attempting to read {js['TEST_DATA']}"))
    try:
        file = open(js['TEST_DATA'], 'r')
        reader = csv.reader(file)
        counter = 0
        file.close()

        for record in reader:
            cursor.execute(f"INSERT INTO {TABLE_NAME} (name, service, email, contactNumber, responded) VALUES (%s,%s,%s,%s,%s)", (f"{record[0]}", f"{record[1]}", f"{record[2]}", f"{record[3]}", f"0"))
            db.commit()
            counter += 1

        logging.debug('{}'.format(f"Successfully read {counter} files into {TABLE_NAME}"))
    except FileNotFoundError as e:
        logging.debug('{}'.format(f"ERROR: {e.errno} - MESSAGE: {e.strerror}"))
    except mysql.connector.Error as e:
        logging.debug('{}'.format(f"ERROR: {e.errno} - SQLSTATE value: {e.sqlstate} - Error Message: {e.msg}"))


def open_outlook():
    try:
        logging.debug('{}'.format(f"Attempting to open Outlook Application"))
        os.startfile(OUTLOOK_LOC)
        logging.debug('{}'.format(f"Successfully opened Outlook Application"))
        os.close
    except FileNotFoundError as e:
        logging.debug('{}'.format(f"ERROR: {e.errno} - {e}"))
    except PermissionError as e:
        logging.debug('{}'.format(f"ERROR: {e.errno} - {e}"))

## MAIN CODE
logging.basicConfig(filename='Logs.log', level=logging.DEBUG, format='%(asctime)s:%(message)s')

# Reading user configuration
with open('information.txt') as f:
    data = f.read()
    js = json.loads(data)
    f.close()

# Constants
HOST = js["HOST"]
USER = js["USER"]
PASSWD = js["PASSWD"]
DATABASE = js["DATABASE"]
TABLE_NAME = js["TABLE_NAME"]
APP_NAME = js["APP_NAME"]
OUTLOOK_LOC = js["OUTLOOK_LOC"]
MAX_RESULTS_PPAGE = js["MAX_RESULTS_PPAGE"]

# Reading available services
services = []
with open('services.txt') as f:
    services = [l for l in (line.strip() for line in f) if l]
    services.insert(0, "None")
    f.close()


db = None
while (db == None):
    db = connect_database()


cursor = db.cursor(buffered=True)

if js["READ_TEST_DATA"] == "True":
    read_test_data()

if js["OPEN_OUTLOOK"] == "True":
    open_outlook()

app = App()
app.mainloop()

os.system('taskkill /F /IM outlook.exe')