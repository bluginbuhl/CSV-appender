from openpyxl import load_workbook
import csv, os, sys
from tkinter import filedialog as fd
import tkinter as tk


application_window = tk.Tk()
application_window.withdraw()

ftypes = [('files', '.txt .csv .xls .xlsx'),
          ('csv files', '.csv'),
          ('Excel workbooks', '.xls')]

csv_files = fd.askopenfilenames(parent=application_window,
                                initialdir=os.getcwd(),
                                title="Please select one or more csv files to load:",
                                filetypes=ftypes)

if csv_files == '':
    print('No data chosen for import. Exiting.')
    sys.exit()

dest_filename = fd.askopenfilename(parent=application_window,
                                   initialdir=os.getcwd(),
                                   title="Please select the file to write to:",
                                   filetypes=ftypes)

if dest_filename == '':
    print('No Excel file chosen for data. Exiting')
    sys.exit()

save_name = fd.asksaveasfilename(parent=application_window,
                                 initialdir=os.getcwd(),
                                 title="Choose a name for the modified file:",
                                 filetypes=ftypes)

if save_name == '':
    print('No file name chosen for saving. Exiting')
    sys.exit()


wb = load_workbook(dest_filename)
sheet = wb.active


for file in csv_files:
    csv_file_name = os.path.basename(file)
    out_file_name = os.path.basename(dest_filename)
    with open(file) as csv_file:
        print(f'Loading data from {csv_file_name}')
        csv_reader = csv.reader(csv_file, delimiter=";")
        next(csv_reader, None)
        i = 0
        for row in csv_reader:
            sheet.append(row)
            i += 1
        print(f'\n\tAppended {i} rows of data from'
              '\n\t\t{csv_file_name} to {out_file_name}')


wb.save(filename=save_name)
print(f'Saved data to {os.path.basename(save_name)}')