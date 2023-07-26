import customtkinter as ctk
from tkinter import messagebox
import win32com.client
import mysql.connector
from functools import partial
import os
import logging
import csv


class App(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("1250x720")
        self.title("Perry Gardiner Database")

        self.GLOBAL_RESULTS = []
        self.MAX_PAGES = 0
        self.CURRENT_PAGE = 1
        self.RESULTS_PER_PAGE = 25

        # Setting window appearances
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        # Clears result frame when update is needed
        def clear_frame():
            for widget in frameRightResults.winfo_children():
                widget.destroy()

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
            results = cursor.fetchall()

            # Calculate max pages and configure page 1
            if len(results) % self.RESULTS_PER_PAGE != 0:
                self.MAX_PAGES = (len(results) // self.RESULTS_PER_PAGE) + 1
            else:
                self.MAX_PAGES = (len(results) // self.RESULTS_PER_PAGE)
            currentPage.configure(text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
            
            # Split results into array of lists for each page
            self.GLOBAL_RESULTS = [results[x:x+self.RESULTS_PER_PAGE] for x in range(0, len(results), self.RESULTS_PER_PAGE)]

            # Clear frame and load Page 1
            if len(self.GLOBAL_RESULTS) == 0:
                clear_frame()
                return
            
            clear_frame()
            load_results(self.GLOBAL_RESULTS[0])
        

        def button_event_delete(id, x):
            # Selects individual from database, deletes and logs
            cursor.execute(f"SELECT * FROM {TABLE_NAME} WHERE personID={id}")
            person = cursor.fetchone()
            cursor.execute(f"DELETE FROM {TABLE_NAME} WHERE personID={id}")
            db.commit()
            logging.debug('{}'.format(f"Successfully deleted entry from: {TABLE_NAME} - INFO: {person[0], person[1], person[2], person[3], person[4]}"))

            # Removes individual from global list and reloads frame
            self.GLOBAL_RESULTS[self.CURRENT_PAGE - 1].pop(x)
            clear_frame()
            load_results(self.GLOBAL_RESULTS[self.CURRENT_PAGE - 1])
            return


        def button_event_email_open(emailAddress):
            outlook = win32com.client.Dispatch('Outlook.Application')
            email = outlook.CreateItem(0)
            email.To = emailAddress
            email.Display(True)
            return


        def button_event_reload_services():
            # Reading available services
            with open('services.txt') as f:
                services = f.read().splitlines()
                services.insert(0, "None")
                f.close()
            
            serviceEntry.configure(values=services)
            serviceSearch.configure(values=services)
            return
        
        
        def button_event_edit_services():
            os.system('services.txt')
            return
        
        
        def button_event_page_down():
            if self.CURRENT_PAGE > 1:
                self.CURRENT_PAGE -= 1
                currentPage.configure(text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
                clear_frame()
                load_results(self.GLOBAL_RESULTS[self.CURRENT_PAGE - 1])
            return
        

        def button_event_page_up():
            if self.CURRENT_PAGE < self.MAX_PAGES:
                self.CURRENT_PAGE += 1
                currentPage.configure(text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
                clear_frame()
                load_results(self.GLOBAL_RESULTS[self.CURRENT_PAGE - 1])
            return
        
        def slider_event(value):
            self.RESULTS_PER_PAGE = int(value)
            resultsShow.configure(text=int(value))


        def load_results(results):
            for x in range(len(results)):
                resultFrame = ctk.CTkFrame(master=frameRightResults)
                resultFrame.pack(padx=5, pady=3, anchor=ctk.W, fill=ctk.BOTH, expand=True)

                # Results
                ctk.CTkLabel(master=resultFrame, corner_radius=0, text=results[x][0], width=150).grid(row=x, column=1, padx=10, pady=5)
                ctk.CTkLabel(master=resultFrame, corner_radius=0, text=results[x][1], width=150).grid(row=x, column=2, padx=10, pady=5)
                ctk.CTkLabel(master=resultFrame, corner_radius=0, text=results[x][2], width=200).grid(row=x, column=3, padx=10, pady=5)
                ctk.CTkLabel(master=resultFrame, corner_radius=0, text=results[x][3], width=150).grid(row=x, column=4, padx=10, pady=5)

                # Delete and Open Email Buttons
                ctk.CTkButton(master=resultFrame, text="Delete", width=70, command=partial(button_event_delete, results[x][5], x)).grid(row=x, column=5, padx=2, pady=5, sticky=ctk.E)
                ctk.CTkButton(master=resultFrame, text="Open Email", width=80, command=partial(button_event_email_open, results[x][2])).grid(row=x, column=6, padx=2, pady=5, sticky=ctk.E)


        ## RESULTS FRAME
        frameRight = ctk.CTkFrame(master=self)
        frameRight.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True, side=ctk.RIGHT)

        # Title
        ctk.CTkLabel(master=frameRight, text="RESULTS", fg_color="transparent", font=("Barlow Condensed", 25)).pack(pady=10)

        frameRightResults = ctk.CTkScrollableFrame(master=frameRight)
        frameRightResults.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True, anchor=ctk.S)

        resultTogglesFrame = ctk.CTkFrame(master=frameRight)
        resultTogglesFrame.pack(padx=10, pady=10, fill=None, expand=False)
        pageSelectionFrame = ctk.CTkFrame(master=resultTogglesFrame)
        pageSelectionFrame.pack(padx=10, pady=10, fill=None, expand=False, side=ctk.LEFT)
        ctk.CTkButton(master=pageSelectionFrame, text="<", command=button_event_page_down).pack(padx=10, pady=10, side=ctk.LEFT)
        currentPage = ctk.CTkLabel(master=pageSelectionFrame, text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
        currentPage.pack(padx=10, pady=10, side=ctk.LEFT)
        ctk.CTkButton(master=pageSelectionFrame, text=">", command=button_event_page_up).pack(padx=10, pady=10, side=ctk.LEFT)

        frameRightChild3 = ctk.CTkFrame(master=resultTogglesFrame)
        frameRightChild3.pack(padx=10, pady=10, fill=None, expand=False, side=ctk.LEFT)

        resultsShow = ctk.CTkLabel(master=frameRightChild3, text=self.RESULTS_PER_PAGE)
        resultsShow.pack()
        resultsSlider = ctk.CTkSlider(master=frameRightChild3, from_=10, to=50, number_of_steps=40, command=slider_event)
        resultsSlider.pack(padx=10, pady=10)
        resultsSlider.set(self.RESULTS_PER_PAGE)


        ## NEW ENTRIES FRAME
        frameUpperLeft = ctk.CTkFrame(master=self)
        frameUpperLeft.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True)

        # Title
        ctk.CTkLabel(master=frameUpperLeft, text="NEW ENTRIES", fg_color="transparent", font=("Barlow Condensed", 25)).pack(pady=10)

        nameEntry = ctk.CTkEntry(master=frameUpperLeft, placeholder_text="Name", width=200)
        nameEntry.pack(pady=5, padx=20)

        serviceEntry = ctk.CTkOptionMenu(master=frameUpperLeft, values=services, width=200)
        serviceEntry.set("Select Service")
        serviceEntry.pack(pady=5, padx=20)

        emailEntry = ctk.CTkEntry(master=frameUpperLeft, placeholder_text="Email", width=200)
        emailEntry.pack(pady=5, padx=20)

        contactEntry = ctk.CTkEntry(master=frameUpperLeft, placeholder_text="Contact Number", width=200)
        contactEntry.pack(pady=5, padx=20)

        respondedCheckBox = ctk.CTkCheckBox(master=frameUpperLeft, text="Responded?")
        respondedCheckBox.pack(pady=5, padx=20)

        # Add Button
        ctk.CTkButton(master=frameUpperLeft, text="ADD", command=button_event_add).pack(side="bottom", pady=10, padx=20)


        ## SEARCH FRAME
        searchFrame = ctk.CTkFrame(master=self)
        searchFrame.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True)

        # Title
        ctk.CTkLabel(master=searchFrame, text="SEARCH", fg_color="transparent", font=("Barlow Condensed", 25)).pack(pady=10)

        serviceSearch = ctk.CTkOptionMenu(master=searchFrame, values=services, width=200)
        serviceSearch.set("Search by Service")
        serviceSearch.pack(pady=5, padx=20)

        nameSearch = ctk.CTkEntry(master=searchFrame, placeholder_text="Search by Name", width=200)
        nameSearch.pack(pady=5, padx=20)

        respondedSearch = ctk.CTkCheckBox(master=searchFrame, text="Filter by Responded")
        respondedSearch.pack(pady=5, padx=10)

        alphabeticalCheck = ctk.CTkCheckBox(master=searchFrame, text="Order Alphabetically")
        alphabeticalCheck.pack(padx=5, pady=10)

        # Search Button
        ctk.CTkButton(master=searchFrame, text="SEARCH", command=button_event_search).pack(pady=10, padx=20, side=ctk.BOTTOM)


        ## SERVICES FRAME
        servicesFrame = ctk.CTkFrame(master=self)
        servicesFrame.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True)

        # Title
        ctk.CTkLabel(master=servicesFrame, text="SERVICES LIST", fg_color="transparent", font=("Barlow Condensed", 25)).pack(pady=10)

        # Reload and Edit services button
        ctk.CTkButton(master=servicesFrame, text="RELOAD SERVICES LIST", command=button_event_reload_services).pack(pady=5, padx=20, side=ctk.BOTTOM)
        ctk.CTkButton(master=servicesFrame, text="EDIT SERVICES LIST", command=button_event_edit_services).pack(pady=5, padx=20, side=ctk.BOTTOM)


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



logging.basicConfig(filename='Logs.log', level=logging.DEBUG, format='%(asctime)s:%(message)s')

## MAIN CODE
# Reading available services
with open('information.txt') as f:
    data = f.read().splitlines()
    f.close()

# Constants
HOST = data[0]
USER = data[1]
PASSWD = data[2]
DATABASE = data[3]
TABLE_NAME = data[4]

# Reading available services

with open('services.txt') as f:
    services = f.read().splitlines()
    services.insert(0, "None")
    f.close()

db = None
while (db == None):
    db = connect_database()


cursor = db.cursor(buffered=True)

#myFile = open('testData.csv', 'r')
#reader = csv.reader(myFile)
#for record in reader:
#    cursor.execute(f"INSERT INTO {TABLE_NAME} (name, service, email, contactNumber, responded) VALUES (%s,%s,%s,%s,%s)", (f"{record[0]}", f"{record[1]}", f"{record[2]}", f"{record[3]}", f"0"))
#    db.commit()



app = App()
app.mainloop()
