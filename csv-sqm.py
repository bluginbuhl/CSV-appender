#!/usr/bin/env python

import csv
import os
import sys
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
import tkinter as tk

# load the tk interface
application_window = tk.Tk()  # define the window as a Tk class
application_window.withdraw()  # withdraw the root window

# define allowed filetypes for dialoges
ftypes = [('files', '.txt .csv .xls .xlsx'),
          ('csv files', '.csv'),
          ('Excel workbooks', '.xls')]


# ask the user to select one or more data files to load
csv_files = filedialog.askopenfilenames(parent=application_window,
                                        initialdir=os.getcwd(),
                                        title="Please select one or more csv" +
                                              " files to load:",
                                        filetypes=ftypes)


# checks to see if data has been selected. Exits if none.
if csv_files == '':
    print("No data chosen for import. Exiting.")
    sys.exit()

# prompt user for the destination file to append the csv data
dest_filename = filedialog.askopenfilename(parent=application_window,
                                           initialdir=os.getcwd(),
                                           title="Please select the file to" +
                                                 " write to:",
                                           filetypes=ftypes)

short_dest_name = os.path.basename(dest_filename)

# checks to see that a destination file is chosen
if dest_filename == '':
    print("No destination file chosen. Exiting")
    sys.exit()


# prompt user for a file name to save to
save_name = filedialog.asksaveasfilename(parent=application_window,
                                         initialdir=os.getcwd(),
                                         title="Choose a name for the modified"
                                               + " file:",
                                         filetypes=ftypes)

# if no save name is chosen, it will overwrite the destination file
if save_name == '':
    answer = messagebox.askokcancel("Warning:", "The file you are appending to"
                                    + " will be overwritten! Are you sure you"
                                    + " want to continue?\n\n"
                                    + f'File: "{short_dest_name}" will be'
                                    + " overwritten.")
    if answer is True:
        save_name = dest_filename
    else:
        print("No data saved.")
        sys.exit()

# loads the workbook object if it exists
if dest_filename:
    wb = load_workbook(dest_filename)
else:
    print("Program exited (no destination file chosen).")

# will load the correct data sheet from the selected workbook
try:
    sheet = wb['SQM']  # looks for SQM sheet
except KeyError:
    print("Error: the selected workbook has no sheet named 'SQM'")


# iterate through the selected CSV files
"""Currently, this doesn't format the data correctly. Also, it overwrites \
charts on other sheets besides the active one.
"""
for file in csv_files:
    csv_file_name = os.path.basename(file)
    out_file_name = os.path.basename(dest_filename)
    with open(file) as csv_file:
        print(f'Loading data from {csv_file_name}')

        # parse the data from the csv file
        csv_reader = csv.reader(csv_file, delimiter=";")
        next(csv_reader, None)  # skips the first row (header)
        i = 0
        for row in csv_reader:
            sheet.append(row)
            i += 1
        print(f'\tAppended {i} rows of data from "{csv_file_name}" to' +
              f'"{out_file_name} ({sheet})"\n')

# saves the workbook - CAUTION: no warning for overwrite!
wb.save(filename=save_name)
print(f'Saved data to {os.path.basename(save_name)}')
