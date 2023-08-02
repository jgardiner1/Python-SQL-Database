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
# implement select all entries on page
# play around with background colours of frames
# reset scrollbar when navigating pages or performing another query

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
            curPage.configure(text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
            
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
                cursor.execute(f"INSERT INTO {TABLE_NAME} (name, service, email, contactNumber, responded) VALUES (%s,%s,%s,%s,%s)", (f"{nameEntry.get()}", f"{serviceEntry.get()}", f"{emailEntry.get()}", f"{contactEntry.get()}", f"{respondedEntry.get()}"))
                db.commit()
                logging.debug('{}'.format(f"Successfully insert into: {TABLE_NAME} - INFO: {nameEntry.get()}, {serviceEntry.get()}, {emailEntry.get()}, {contactEntry.get()}, {respondedEntry.get()}"))
            except mysql.connector.Error as e:
                logging.debug('{}'.format(f"ERROR: {e.errno} - SQLSTATE value: {e.sqlstate} - Error Message: {e.msg}"))
            
            # resetting all fields
            nameEntry.delete(0, 50)
            serviceEntry.set("None")
            emailEntry.delete(0, 50)
            contactEntry.delete(0, 50)
            respondedEntry.deselect()


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
            
            if alphabeticalSearch.get() == 1:
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
            curPage.configure(text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
            
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
                curPage.configure(text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
                clear_frame()
                load_results(self.CURRENT_PAGE - 1)
            return
        

        def button_event_page_up():
            if self.CURRENT_PAGE < self.MAX_PAGES:
                self.CURRENT_PAGE += 1
                curPage.configure(text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
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
        rightFr = ctk.CTkFrame(master=self)
        rightFr.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True, side=ctk.RIGHT)

        # Title
        ctk.CTkLabel(master=rightFr, text="RESULTS", fg_color="transparent", font=("Barlow Condensed", 25)).pack(pady=7)

        # Frame holds results
        frameResults = ctk.CTkScrollableFrame(master=rightFr)
        frameResults.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True)

        # Frame for under main results frame. Stores page selection and results per page
        rightBottomFr = ctk.CTkFrame(master=rightFr, fg_color="gray13")
        rightBottomFr.pack(padx=10, pady=10, fill=ctk.X, expand=True, side=ctk.LEFT)

        pageNavFr = ctk.CTkFrame(master=rightFr)
        pageNavFr.pack(padx=10, pady=10, fill=ctk.X, expand=True, side=ctk.LEFT)

        rightBottomFr3 = ctk.CTkFrame(master=rightFr, height=20)
        rightBottomFr3.pack(padx=10, pady=10, fill=ctk.X, expand=True, side=ctk.LEFT)

        ctk.CTkButton(master=rightBottomFr, text="DELETE SELECTED RESULTS", command=button_event_delete).pack(padx=10, pady=10, side=ctk.LEFT)

        # Page Selection frame to store buttons and current page info
        pageSelectFr = ctk.CTkFrame(master=pageNavFr)
        pageSelectFr.pack(padx=10, pady=10, fill=None, expand=True, side=ctk.LEFT, anchor=ctk.CENTER)
        ctk.CTkButton(master=pageSelectFr, text="<", command=button_event_page_down).pack(padx=10, pady=10, side=ctk.LEFT)
        curPage = ctk.CTkLabel(master=pageSelectFr, text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
        curPage.pack(padx=10, pady=10, side=ctk.LEFT)
        ctk.CTkButton(master=pageSelectFr, text=">", command=button_event_page_up).pack(padx=10, pady=10, side=ctk.LEFT)


        ## SEARCH FRAME
        middleLeftFr = ctk.CTkFrame(master=self)
        middleLeftFr.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True)

        # Title
        ctk.CTkLabel(master=middleLeftFr, text="SEARCH", fg_color="transparent", font=("Barlow Condensed", 25)).pack(pady=7)

        serviceSearch = ctk.CTkOptionMenu(master=middleLeftFr, values=services, width=200)
        serviceSearch.set("Search by Service")
        serviceSearch.pack(pady=5, padx=20)

        nameSearch = ctk.CTkEntry(master=middleLeftFr, placeholder_text="Search by Name", width=200)
        nameSearch.pack(pady=5, padx=20)

        respondedSearch = ctk.CTkCheckBox(master=middleLeftFr, text="Filter by Responded")
        respondedSearch.pack(pady=5, padx=10)

        alphabeticalSearch = ctk.CTkCheckBox(master=middleLeftFr, text="Order Alphabetically")
        alphabeticalSearch.pack(padx=5, pady=10)

        # Search Button
        ctk.CTkButton(master=middleLeftFr, text="SEARCH", command=button_event_search).pack(pady=10, padx=20, side=ctk.BOTTOM)


        ## NEW ENTRIES FRAME
        topLeftFr = ctk.CTkFrame(master=self)
        topLeftFr.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True)

        # Title
        ctk.CTkLabel(master=topLeftFr, text="NEW ENTRIES", fg_color="transparent", font=("Barlow Condensed", 25)).pack(pady=7)

        nameEntry = ctk.CTkEntry(master=topLeftFr, placeholder_text="Name", width=200)
        nameEntry.pack(pady=5, padx=20)

        serviceEntry = ctk.CTkOptionMenu(master=topLeftFr, values=services, width=200)
        serviceEntry.set("Select Service")
        serviceEntry.pack(pady=5, padx=20)

        emailEntry = ctk.CTkEntry(master=topLeftFr, placeholder_text="Email", width=200)
        emailEntry.pack(pady=5, padx=20)

        contactEntry = ctk.CTkEntry(master=topLeftFr, placeholder_text="Contact Number", width=200)
        contactEntry.pack(pady=5, padx=20)

        respondedEntry = ctk.CTkCheckBox(master=topLeftFr, text="Responded?")
        respondedEntry.pack(pady=5, padx=20)

        # Add Button
        ctk.CTkButton(master=topLeftFr, text="ADD", command=button_event_add).pack(side="bottom", pady=10, padx=20)


        ## SERVICES FRAME
        bottomLeftFr = ctk.CTkFrame(master=self)
        bottomLeftFr.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True)

        # Title
        ctk.CTkLabel(master=bottomLeftFr, text="EDIT SERVICES", fg_color="transparent", font=("Barlow Condensed", 25)).pack(pady=7)

        addServiceFr = ctk.CTkFrame(master=bottomLeftFr, fg_color="gray13")
        addServiceFr.pack(padx=5, pady=5, fill=ctk.BOTH, expand=True)
        addServiceEntry = ctk.CTkEntry(master=addServiceFr, placeholder_text="New Service")
        addServiceEntry.pack(padx=5, pady=5, side=ctk.LEFT)
        ctk.CTkButton(master=addServiceFr, text="ADD", command=button_event_add_service).pack(padx=5, pady=5, side=ctk.LEFT)
        
        # Reload and Edit services button
        editServiceFr = ctk.CTkFrame(master=bottomLeftFr, fg_color="gray13")
        editServiceFr.pack(padx=5, pady=5, fill=ctk.Y, expand=True)
        ctk.CTkButton(master=editServiceFr, text="RELOAD SERVICES", command=button_event_reload_services).pack(padx=5, pady=5, side=ctk.BOTTOM)
        ctk.CTkButton(master=editServiceFr, text="EDIT SERVICES", command=button_event_edit_services).pack(padx=5, pady=5, side=ctk.BOTTOM)


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