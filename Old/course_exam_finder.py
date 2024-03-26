from ttkbootstrap.constants import *
import ttkbootstrap as ttk
from glob import glob
import csv
import os

def resource_path(relative_path):
            #Get absolute path to resource, works for dev and for PyInstaller
    try:
        base_path = os.path.abspath(".") + "\\"
        file_path = glob(base_path + relative_path)
        return file_path[0]
    except Exception:
        base_path = os.path.abspath(".") + "\\CEF\\"
        file_path = glob(base_path + relative_path)
        return file_path[0]

class App(ttk.Window):
    def __init__(self):
        super().__init__()

        self.classes = ""
        self.title("Course Exam Finder")
        self.resizable(False,False)
        self.geometry(f"650x180+{1920//2-650//2}+360")
        style = ttk.Style("darkly")
        style.configure("TButton", font=(None, 14))

        label = ttk.Label(self, text="Course Exam Finder", foreground="white", font=(None, 20))
        label.pack(padx=15, pady=15)

        self.entry = ttk.Entry(self, bootstyle=(LIGHT), font=(None, 18))
        self.entry.pack(padx=15, ipady=2, fill="x")

        button = ttk.Button(self, text="Find", width=6, bootstyle=INFO, command=self.read_create)
        button.pack(padx=20, pady=15, ipady=2)

        self.bind("<Return>", lambda event: self.read_create())

    def read_create(self):
        self.classes = self.entry.get()
        file_path = resource_path("*schedule.csv")
        with open(file_path) as file:
            csv_reader = csv.reader(file)        

            with open("My exams.csv", mode="w", newline="") as file:
                csv_writer = csv.writer(file)
                csv_writer.writerow(["Course Exam","Start Time","Classrooms"])

                if "Final" in file_path:
                    for row in csv_reader:
                        if row[0][:8].replace(" ","").upper() in self.classes.replace(" ","").upper():
                            csv_writer.writerow((row[0], row[1]+" "+row[2][:5], row[3]))
                else:
                    for row in csv_reader:
                        if row[0][:8].replace(" ","").upper() in self.classes.replace(" ","").upper():
                            csv_writer.writerow((row[0], row[1], row[3]))

app = App()
app.mainloop()
