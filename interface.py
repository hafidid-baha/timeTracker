from tkinter import *
from tkinter import messagebox
from database import Database
from tkinter import filedialog
import sys
import os
from workBook import WorkBook
import time


class Ui:
    def __init__(self):
        self.root = Tk()
        self.root.title("Time Trucker")
        self.root.resizable(False, False)
        # show windows on the screen center
        window_height = 100
        window_width = 400
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x_cordinate = int((screen_width / 2) - (window_width / 2))
        y_cordinate = int((screen_height / 2) - (window_height / 2))
        self.root.geometry("{}x{}+{}+{}".format(window_width, window_height, x_cordinate, y_cordinate))
        # create a text entry
        self.e = Entry(self.root, bg="#eee", fg="black")
        self.e.grid(row=0, sticky='we', column=0, columnspan=2, ipady=10, ipadx=10)
        self.e.insert(0, Database.get_last_task()[0][1])
        self.e.focus()
        # create save task btn and export btn
        mbtn = Button(self.root, text="Save Task", padx=10, pady=5, command=self.save_task)
        mbtn.grid(row=1, column=0, sticky="W", pady=10)
        export_btn = Button(self.root, text="Export Excel File", padx=10, pady=5, command=self.export_file)
        export_btn.grid(row=1, column=1, sticky="W", pady=10)

        self.root.grid_columnconfigure(0, weight=1)
        # show the window always on the top of all author apps
        self.root.attributes('-topmost', True)
        self.root.mainloop()

    def save_task(self):
        database = Database()
        database.create_task(self.e.get())
        self.root.wm_state('withdrawn')
        # show the app main window again every 1 hour
        while self.root.wm_state() == "withdrawn":
            time.sleep(3600)
            if self.root.wm_state() == "withdrawn":
                self.root.wm_state("normal")

    def export_file(self):
        filename = filedialog.askdirectory()
        WorkBook(os.path.join(filename, "timeTracker.xlsx"))
        messagebox.showinfo(title="Export Success", message="Your Excel File Has Been Saved")


sys.path.append(".")
