# Tool created for speeding up the process of tidying up user exports from MS365

# Imports
import pandas as pd
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog

# Initial declaration of global variables.
file = None
df = None
to_remove = None


# File dialog popup to select the input file.
def open_file():
    global file
    file = filedialog.askopenfilename()
    print(file)


# Read input CSV file and create list of columns to be removed.
def tidy_csv():
    global df, to_remove, file
    df = pd.read_csv(file)

    to_remove = ["Block credential", "City", "Country/Region", "Department", "DirSyncEnabled", "Fax", "Last dirsync time",
                 "Last password change time stamp", "License assignment details", "Mobile Phone", "Oath token meta data",
                 "Object Id", "Office", "Password never expires", "Phone number", "Postal code", "Preferred data location",
                 "Preferred language", "Proxy addresses", "Release track", "Soft deletion time stamp", "State",
                 "Street address", "Strong password required", "Title", "Usage location", "When created"]
    write_csv()


# Remove the columns selected above along with removing guests if selected, then write the new CSV.
def write_csv():
    global df, to_remove, df_new, var1
    df.drop(columns=to_remove, inplace=True)
    if var1.get() == 1:
        df = df[df["User principal name"].str.contains("#EXT#") == False]
    else:
        df = df
    new_file = filedialog.asksaveasfile()
    df.to_csv(new_file, index=False)
    messagebox.showinfo(title="Success", message="File tidied successfully!")


# GUI
window = Tk()
var1 = IntVar()
window.title("MS365 Tidy Uperer")
window.config(padx=20, pady=20)

file_button = Button(text="Select CSV File", command=open_file)
file_button.grid(column=0, row=0, padx=10, pady=10)

guest_checkbox = Checkbutton(text="Remove Guests?", variable=var1, onvalue=1, offvalue=0)
guest_checkbox.grid(column=0, row=1, padx=10, pady=10)

convert_button = Button(text="Tidy CSV", command=tidy_csv)
convert_button.grid(column=0, row=2, padx=10, pady=10)

window.mainloop()
