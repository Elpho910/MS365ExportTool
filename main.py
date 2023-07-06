# Tool created for speeding up the process of tidying up user exports from MS365

import pandas as pd
from tkinter import *
from tkinter import messagebox


def tidy_csv():
    df = pd.read_csv("input.csv")

    to_remove = ["Block credential", "City", "Country/Region", "Department", "DirSyncEnabled", "Fax", "Last dirsync time",
                 "Last password change time stamp", "License assignment details", "Mobile Phone", "Oath token meta data",
                 "Object Id", "Office", "Password never expires", "Phone number", "Postal code", "Preferred data location",
                 "Preferred language", "Proxy addresses", "Release track", "Soft deletion time stamp", "State",
                 "Street address", "Strong password required", "Title", "Usage location", "When created"]

    df.drop(columns=to_remove, inplace=True)
    df = df[df["User principal name"].str.contains("#EXT#") == False]
    df.to_csv("output.csv", index=False)
    messagebox.showinfo(title="Success", message="File tidied successfully!")


window = Tk()
window.title("MS365 Tidy Uperer")
window.config(padx=20, pady=20)

word_label = Label(text="Ensure file from MS365 is named input.csv")
word_label.grid(column=0, row=0, padx=10, pady=10)

convert_button = Button(text="Tidy CSV", command=tidy_csv)
convert_button.grid(column=1, row=2, padx=10, pady=10)

window.mainloop()
