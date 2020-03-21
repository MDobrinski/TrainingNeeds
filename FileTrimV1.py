import csv
import datetime
import os
import xlsxwriter

import tkinter as tk
from tkinter import filedialog
from tkinter import ttk
from tkinter import Menu
from tkinter import messagebox

import pandas as pd


def get_file():
    global filename
    filename = filedialog.askopenfilename(initialdir="/Users/mdobrinski/data", title="Select A File", filetypes=[
        ("CSV files", "*.csv")])
#   print(os.path.dirname(filename))
    out_path = os.path.dirname(filename) + "\\output"
    if os.path.exists(out_path):
        print("***** The output path exists ********* ")
        os.chdir(out_path)
    else:
        print(out_path, " DOES NOT EXIST!!!")
        os.mkdir(out_path)
        os.chdir(out_path)
    print(os.getcwd())
    print("The output directory is:", out_path)
    l_text.set(filename)
    op_text.set(out_path)
    # print(filename, " in get_file")


def about_message():
    messagebox.showinfo("About", "Application to parse a csv file of learning needs into individual district needs."
                                 "\n\nA folder called \"output\" will be created under the source folder and the "
                                 "generated files will be placed in the output folder."
                                 "\n\n\n\n version 3.2")


def end_prog():
    root.destroy()


def trim_file():
    global filename
    current_date = datetime.date.today().strftime("%m%d%Y")

    # Open data file for reading into dictionary
    data = pd.read_csv(filename, index_col="Item Type")

    # dropping unwanted columns
    data.drop(["Revision Date", "Revision Number", "Active User", "Middle Initial", "Assignment Type ID",
               "Assignment Type", "Required", "Assignment Date", "Required Date", "Expiration Date",
               "Preferred Time zone", "Organization ID", "Organization Description", "Job Location ID",
               "Job Code ID", "Job Code", "Employee Status ID", "Employee Status", "Employee Type ID",
               "Employee Type", "Supervisor Middle Initial"], axis=1, inplace=True, errors='ignore')

    data.to_csv(filename)
    print("File completed!")
    result_text.set("Results: File completed!")


# *************************** Create the Window Objects **************************


root = tk.Tk()  # ************ Main (root) Window
root.option_add('*tearOff', False)
root.title('Trim Training Needs Report')
root.geometry('650x350+200+200')
root.minsize(650, 300)
menubar = Menu(root)
root.config(menu=menubar)
file = Menu(menubar)
help_ = Menu(menubar)
menubar.add_cascade(menu=file, label="File")
menubar.add_cascade(menu=help_, label="Help")
file.add_command(label='Open...', command=lambda: get_file())
file.add_separator()
file.add_command(label='Exit', command=lambda: end_prog())
help_.add_command(label='About', command=lambda: about_message())

filename = "No file"
l_text = tk.StringVar()
result_text = tk.StringVar()
result_text.set("Results: ")
op_text = tk.StringVar()
op_text.set("                   ")
bg_color = '#BED8BC'

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S))
mainframe.config(relief=tk.RIDGE, borderwidth=5, style='TFrame')
op_frame = ttk.Frame(root)
op_frame.grid(column=0, row=1, sticky=(tk.N, tk.W, tk.E, tk.S))
op_frame.config(relief=tk.RIDGE, borderwidth=5, style='TFrame')
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)
style = ttk.Style()
# print(style.theme_names())
# print(style.theme_use())
style.theme_use('vista')
style.configure('TFrame', background=bg_color)
style.configure('TLabel', background=bg_color, font=('Arial', 9))


logo = tk.PhotoImage(file="CJ_Logo_HalfSize.gif")

# debug statement print(filename, " in Main")
# debug statement print("l_text = ", l_text.get)
l_text.set(filename)

open_btn = ttk.Button(mainframe, text="Open File", command=get_file)
open_btn.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
parse_btn = ttk.Button(mainframe, text="Trim File", command=trim_file)
parse_btn.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
quit_btn = ttk.Button(mainframe, text="Exit", command=end_prog)
quit_btn.grid(row=3, column=0, padx=5, pady=5, sticky=tk.W)

fn_label = ttk.Label(mainframe, textvariable=l_text, relief="solid", width=70)
fn_label.grid(row=0, column=1, columnspan=3, sticky='nsew', padx=5, pady=5)
result_lbl = ttk.Label(mainframe, textvariable=result_text, width=25, style='TLabel')
result_lbl.grid(row=1, column=1, columnspan=2, sticky='nsw', padx=5, pady=5)
logo_lbl = ttk.Label(mainframe, image=logo, style='TLabel')
logo_lbl.grid(row=3, column=2, columnspan=2, rowspan=2, sticky='nsew', padx=5, pady=5)
op_label = ttk.Label(op_frame, text="Output Folder", style='TLabel')
op_label.grid(row=0, column=0, sticky='nsw', padx=5, pady=5)
path_label = ttk.Label(op_frame, textvariable=op_text, relief="solid", width=70, style='TLabel')
path_label.grid(row=0, column=1, columnspan=3, sticky='nse', padx=5, pady=5)

root.mainloop()
