# cut copy software for photodeprtment
# 02-22-2024
import shutil
import os
import pandas as pd
import tkinter as tk
import openpyxl
from tkinter import ttk
from tkinter import filedialog
import fnmatch

##################################################################################
# Functions

def load_excel_file():

    global excel_path
    excel_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

def select_picture_folder():
    global picture_path
    picture_path = filedialog.askdirectory()

def select_folder_target():
    global target_path
    target_path = filedialog.askdirectory()

def run_report():
    if not 'excel_path' in globals():
        #print("Please select an Excel file first.")
        return
    
    if not 'picture_path' in globals():
        #print("Please select a picture folder first.")
        return

    if not 'target_path' in globals():
        #print("Please select a target folder first.")
        return

    if cut_var.get():
        operation = 'cut'
    else:
        operation = 'copy'

    df = pd.read_excel(excel_path)
    #row_num = df.shape[0]
    
    for index, row in df.iterrows():
            search_string = row['SKU']
            matches = []
            for root, _, files in os.walk(picture_path):
                for filename in files:
                    if fnmatch.fnmatch(filename, f'*{search_string}*'):
                        matches.append(os.path.join(root, filename))
            if matches:
                #print(f"Found files containing '{search_string}':")
                for match in matches:
                    if operation == 'cut':
                        shutil.move(match, target_path)
                        #print(match)
                    else:
                        shutil.copy(match, target_path)
                        #print(match)

                #print(f"No files found containing '{search_string}'.")



##################################################################################
# Main

root = tk.Tk()
root.title("***The Amazin Christian's*** Cut Copy Software")

root.geometry("450x350")

file_frame = ttk.LabelFrame(root, text="Howdy -.~")
file_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

ttk.Label(file_frame, text="Select the Excel file").place(x=10, y=20)
load_button = ttk.Button(file_frame, text="Select Excel", command=load_excel_file)
load_button.place(x=10, y=40)

ttk.Label(file_frame, text="Select picture folder").place(x=10, y=80)
load_button = ttk.Button(file_frame, text="Select Folder", command=select_picture_folder)
load_button.place(x=10, y=100)

ttk.Label(file_frame, text="Select target folder").place(x=10, y=140)
load_button = ttk.Button(file_frame, text="Select Folder", command=select_folder_target)
load_button.place(x=10, y=160)

ttk.Label(file_frame, text="Copy is set by default").place(x=10, y=200)
cut_var = tk.BooleanVar()
checkbutton = ttk.Checkbutton(file_frame, text="Cut", variable=cut_var, onvalue=True, offvalue=False)
checkbutton.place(x=30, y=220)
cut_var.set(False)

ttk.Label(file_frame, text="Run Cut Copy").place(x=10, y=250)
run_button = ttk.Button(file_frame, text="Run", command=run_report)
run_button.place(x=10, y=270)

workbook = None

root.mainloop()