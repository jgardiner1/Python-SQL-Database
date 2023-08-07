import win32com.client
import customtkinter as ctk
from tkinter import messagebox
from functools import partial
import time
import logging
import mysql.connector
import os

log_format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
logger = logging.getLogger(__name__)
logger.setLevel('DEBUG')
file_handler = logging.FileHandler('Logs.log')
formatter = logging.Formatter(log_format)
file_handler.setFormatter(formatter)

logger.addHandler(file_handler)


class ResultPage(ctk.CTkFrame):
    def __init__(self, master, pageNum, **kwargs):
        super().__init__(master, **kwargs)

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


        for x in range(len(self.GLOBAL_RESULTS[pageNum])):
            result = ctk.CTkFrame(master=master)
            result.pack(padx=5, pady=3, fill=ctk.BOTH, expand=True)
            # Results
            if self.GLOBAL_RESULTS[pageNum][x][5] in self.REMOVAL_LIST:
                temp = ctk.CTkCheckBox(master=result, text=None, width=0, command=partial(checkbox_event_entry_selection, app.GLOBAL_RESULTS[pageNum][x][5]))
                temp.grid(row=x, column=1, padx=10, pady=5)
                temp.select()
            else:
                temp = ctk.CTkCheckBox(master=result, text=None, width=0, command=partial(checkbox_event_entry_selection, app.GLOBAL_RESULTS[pageNum][x][5]))
                temp.grid(row=x, column=1, padx=10, pady=5)
                self.CHECK_BOXES.append(temp)
            
            ctk.CTkLabel(master=result, corner_radius=0, text=self.GLOBAL_RESULTS[pageNum][x][0], width=150).grid(row=x, column=2, padx=10, pady=5)
            ctk.CTkLabel(master=result, corner_radius=0, text=self.GLOBAL_RESULTS[pageNum][x][1], width=150).grid(row=x, column=3, padx=10, pady=5)
            ctk.CTkLabel(master=result, corner_radius=0, text=self.GLOBAL_RESULTS[pageNum][x][2], width=200).grid(row=x, column=4, padx=10, pady=5)
            ctk.CTkLabel(master=result, corner_radius=0, text=self.GLOBAL_RESULTS[pageNum][x][3], width=150).grid(row=x, column=5, padx=10, pady=5)

            #Delete and Open Email Buttons
            ctk.CTkButton(master=result, text="Open Email", width=80, command=partial(button_event_email_open, self.GLOBAL_RESULTS[pageNum][x][2])).grid(row=x, column=6, padx=2, pady=5, sticky=ctk.E)


class App(ctk.CTk):
    def __init__(self, APP_NAME, TABLE_NAME, MAX_RESULTS_PPAGE, cursor, db, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("1280x720")
        self.maxsize(1280, 720)
        self.title(APP_NAME)
        self.GLOBAL_RESULTS = []
        self.REMOVAL_LIST = []
        self.MAX_PAGES = 0
        self.CURRENT_PAGE = 1
        self.RESULTS_PER_PAGE = MAX_RESULTS_PPAGE
        self.LAST_QUERY = ""
        self.CHECK_BOXES = []
        self.temp = []

        # Setting window appearances
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        # Reading available services
        services = []
        with open('services.txt') as f:
            services = [l for l in (line.strip() for line in f) if l]
            services.insert(0, "None")
            f.close()

        # Clears result frame when update is needed
        def clear_frame():
            print("\nClearing frame")
            start = time.perf_counter()

            for widget in resultsScroll.winfo_children():
                widget.destroy()

            print("Time to clear frame: ", time.perf_counter()-start)
        

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
                logger.info('{}'.format(f"Successfully insert into: {TABLE_NAME} - INFO: {nameEntry.get()}, {serviceEntry.get()}, {emailEntry.get()}, {contactEntry.get()}, {respondedEntry.get()}"))
            except mysql.connector.Error as e:
                logger.error('{}'.format(f"ERROR: {e.errno} - SQLSTATE value: {e.sqlstate} - Error Message: {e.msg}"))
            
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
                ctk.CTkLabel(master=resultsScroll, text="No Results...").pack()
                return
            
            # Calculate max pages and configure page 1
            if len(results) % self.RESULTS_PER_PAGE != 0:
                self.MAX_PAGES = (len(results) // self.RESULTS_PER_PAGE) + 1
            else:
                self.MAX_PAGES = (len(results) // self.RESULTS_PER_PAGE)
            curPage.configure(text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
            
            # Split results into array of lists for each page
            self.GLOBAL_RESULTS = [results[x:x+self.RESULTS_PER_PAGE] for x in range(0, len(results), self.RESULTS_PER_PAGE)]
            
            load_results(0)
        

        def button_event_delete():
            removal_list = ','.join(str(int(x)) for x in self.REMOVAL_LIST)

            # Selects individual from database, deletes and logs
            try:
                cursor.execute(f"DELETE FROM {TABLE_NAME} WHERE personID IN ({removal_list})")
                db.commit()
                logger.info('{}'.format(f"Successfully deleted {cursor.rowcount} entries from: {TABLE_NAME}"))
                self.REMOVAL_LIST.clear()
                repeat_search()
            except mysql.connector.Error as e:
                logger.error('{}'.format(f"ERROR: {e.errno} - SQLSTATE value: {e.sqlstate} - Error Message: {e.msg}"))
            return

        
        def button_event_add_service():

            if addServiceEntry.get() in button_event_reload_services():
                messagebox.showerror('ERROR', 'Service already within list')
            else:
                file = open('services.txt', 'a')
                file.write(f"\n{addServiceEntry.get()}")
                file.close()

            button_event_reload_services()

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
        

        def checkbox_event_select_all():
            if len(self.CHECK_BOXES) == self.RESULTS_PER_PAGE:
                for checkbox in self.CHECK_BOXES:
                    checkbox.toggle()
                return
        
            for checkbox in self.CHECK_BOXES:
                if checkbox.get() == 0:
                    checkbox.toggle()
            
            self.CHECK_BOXES.clear()

        
        def button_event_page_down():
            if self.CURRENT_PAGE > 1:
                self.CURRENT_PAGE -= 1
                curPage.configure(text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
                selectAllChk.deselect()
                load_results(self.CURRENT_PAGE - 1)
            return


        def button_event_page_up():
            if self.CURRENT_PAGE < self.MAX_PAGES:
                self.CURRENT_PAGE += 1
                curPage.configure(text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
                selectAllChk.deselect()
                load_results(self.CURRENT_PAGE - 1)
            return


        def load_results(pageNum):
            clear_frame()

            print("Loading. Starting counter")
            start = time.perf_counter()

            ResultPage(master=resultsScroll, pageNum=pageNum)
            
            print("Time to load Results: ", time.perf_counter()-start)


        ## RESULTS FRAME
        rightFr = ctk.CTkFrame(master=self)
        rightFr.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True, side=ctk.RIGHT)

        # Title
        ctk.CTkLabel(master=rightFr, text="RESULTS", fg_color="transparent", font=("Barlow Condensed", 25)).pack(pady=7)

        # Frame holds results
        resultsScroll = ctk.CTkScrollableFrame(master=rightFr)
        resultsScroll.pack(padx=10, pady=10, fill=ctk.BOTH, expand=True)

        buttonsFr = ctk.CTkFrame(master=rightFr)
        buttonsFr.pack(fill=ctk.X, expand=True, side=ctk.RIGHT, padx=10, pady=(0, 10))

        # Frame for holding delete results checkBox and button
        leftBottomFr = ctk.CTkFrame(master=buttonsFr, fg_color="gray16")
        leftBottomFr.pack(pady=10, fill=ctk.X, expand=True, side=ctk.LEFT)

        # Frame for under main results frame. Stores page selection and results per page
        pageNavFr = ctk.CTkFrame(master=buttonsFr, fg_color="gray16")
        pageNavFr.pack(pady=10, fill=ctk.X, expand=True, side=ctk.LEFT)

        ctk.CTkFrame(master=buttonsFr, height=20, fg_color="gray16").pack(pady=10, fill=ctk.X, expand=True, side=ctk.LEFT)

        selectAllChk = ctk.CTkCheckBox(master=leftBottomFr, text="SELECT ALL", command=checkbox_event_select_all)
        selectAllChk.pack(padx=(20, 0), anchor=ctk.W)
        ctk.CTkButton(master=leftBottomFr, text="DELETE SELECTED RESULTS", command=button_event_delete).pack(padx=(20, 0), pady=10, anchor=ctk.W)

        # Page Selection frame to store buttons and current page info
        pageSelectFr = ctk.CTkFrame(master=pageNavFr)
        pageSelectFr.pack(padx=10, pady=10, fill=None, expand=True, side=ctk.LEFT, anchor=ctk.CENTER)
        ctk.CTkButton(master=pageSelectFr, text="<", command=button_event_page_down, width=50).pack(padx=10, pady=10, side=ctk.LEFT)
        curPage = ctk.CTkLabel(master=pageSelectFr, text=f"{self.CURRENT_PAGE}/{self.MAX_PAGES}")
        curPage.pack(padx=10, pady=10, side=ctk.LEFT)
        ctk.CTkButton(master=pageSelectFr, text=">", command=button_event_page_up, width=50).pack(padx=10, pady=10, side=ctk.LEFT)


        ## SEARCH FRAME
        middleLeftFr = ctk.CTkFrame(master=self)
        middleLeftFr.pack(padx=10, pady=(10, 0), fill=ctk.BOTH, expand=True)

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
        topLeftFr.pack(padx=10, pady=(10, 0), fill=ctk.BOTH, expand=True)

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
        bottomLeftFr.pack(padx=10, pady=(10, 10), fill=ctk.BOTH, expand=True)

        # Title
        ctk.CTkLabel(master=bottomLeftFr, text="EDIT SERVICES", fg_color="transparent", font=("Barlow Condensed", 25)).pack(pady=7)

        addServiceFr = ctk.CTkFrame(master=bottomLeftFr, fg_color="gray13")
        addServiceFr.pack(padx=5, pady=5, fill=ctk.BOTH, expand=True)

        childFr = ctk.CTkFrame(master=addServiceFr)
        childFr.pack()

        addServiceEntry = ctk.CTkEntry(master=childFr, placeholder_text="New Service")
        addServiceEntry.pack(padx=5, side=ctk.LEFT)
        ctk.CTkButton(master=childFr, text="ADD", command=button_event_add_service).pack(padx=5, side=ctk.LEFT)
        
        # Reload and Edit services button
        editServiceFr = ctk.CTkFrame(master=bottomLeftFr, fg_color="gray13")
        editServiceFr.pack(padx=5, fill=ctk.Y, expand=True)
        ctk.CTkButton(master=editServiceFr, text="RELOAD SERVICES", command=button_event_reload_services).pack(padx=5, pady=5, side=ctk.BOTTOM)
        ctk.CTkButton(master=editServiceFr, text="EDIT SERVICES", command=button_event_edit_services).pack(padx=5, pady=5, side=ctk.BOTTOM)
