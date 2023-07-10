# Tool created for speeding up the process of tidying up user exports from MS365

# Imports
import pandas as pd
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog


class CSVEditor:
    def __init__(self, window):
        self.window = window
        self.var1 = IntVar()
        window.title("MS365 Tidy Uperer")
        window.config(padx=20, pady=20)

        self.file_button = Button(text="Select CSV File", command=self.open_file)
        self.file_button.grid(column=0, row=0, padx=10, pady=10)

        self.guest_checkbox = Checkbutton(text="Remove Guests?", variable=self.var1, onvalue=1, offvalue=0)
        self.guest_checkbox.grid(column=2, row=0, padx=10, pady=10)

        self.convert_button = Button(text="Tidy CSV", command=self.write_csv)
        self.convert_button.grid(column=1, row=0, padx=10, pady=10)

        self.checkboxes = []
        self.checkbox_count = 0
        self.checkbox_vars = []
        self.df = None

    # File dialog popup to select the input file.
    def open_file(self):
        file = filedialog.askopenfilename(filetypes=[('CSV Files', '*.csv')])
        self.df = pd.read_csv(file)

        # Clear old checkboxes
        for checkbox in self.checkboxes:
            checkbox.destroy()
        self.checkboxes.clear()
        self.checkbox_vars.clear()
        self.checkbox_count = 0

        # Add new checkboxes
        for column in self.df.columns:
            self.checkbox_vars.append(IntVar(value=1))
            self.checkboxes.append(Checkbutton(text=f"Remove {column}?", variable=self.checkbox_vars[self.checkbox_count], onvalue=1, offvalue=0))
            if column == "User principal name" or column == "Licenses" or column == "Last name" or column == "First name" or column == "Display name":
                self.checkbox_vars[self.checkbox_count].set(0)
            else:
                self.checkbox_vars[self.checkbox_count].set(1)
            if self.checkbox_count % 2 == 0:
                self.checkboxes[self.checkbox_count].grid(row=self.checkbox_count+1, column=0, sticky=W)
            else:
                self.checkboxes[self.checkbox_count].grid(row=self.checkbox_count, column=1, sticky=W)

            self.checkbox_count += 1

    # Remove the columns selected above along with removing guests if selected, then write the new CSV.
    def write_csv(self):
        if self.df is None:
            messagebox.showerror("No File", "You forgot to load a CSV file!")

        to_remove = [column for column, var in zip(self.df.columns, self.checkbox_vars) if var.get() == 1]

        self.df.drop(columns=to_remove, inplace=True)
        if self.var1.get() == 1:
            self.df = self.df[self.df["User principal name"].str.contains("#EXT#") == False]
        else:
            self.df = self.df
        new_file = filedialog.asksaveasfile(defaultextension=".csv")
        self.df.to_csv(new_file, index=False)
        messagebox.showinfo(title="Success", message="File tidied successfully!")


# GUI
window = Tk()
editor = CSVEditor(window)
window.mainloop()
