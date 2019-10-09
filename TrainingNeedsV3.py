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


def parse_file():
    f_name_list = []
    current_date = datetime.date.today().strftime("%m%d%Y")

    # Open data file for reading into dictionary
    with open(filename, 'r') as csv_file:
        csv_reader = csv.DictReader(csv_file)

        districts = {'Charlotte, TX (TX006)' : '_CTX',
                     'Greeley, CO (CO008)' : '_GCO',
                     'Midland, TX-CH (TX012)' : '_MTX',
                     'San Angelo, TX (TX020)' : '_SATX',
                     'Shafter, CA (CA001)' : '_SCA',
                     'Signal Hill, CA (CA002)' : '_SHCA',
                     'Weatherford, OK (OK004)' : '_WOK',
                     'Williston, ND - (Wireline) (ND009)' : '_WND',
                     'Williston, ND-CH (ND002)' : '_WND2'}

        for district, loc in districts.items():

            # Filter each locations training needs
            f_name = current_date + loc + "_Learning_Needs" + ".csv"
            f_name_list.append(f_name)

            with open(f_name, 'w') as new_file:
                fieldnames = ['Item Type', 'Item ID', 'Item Title', 'User ID', 'Last Name', 'First Name',
                              'Days Remaining', 'Job Location', 'Supervisor ID', 'Supervisor Last Name',
                              'Supervisor First Name']
                csv_writer = csv.DictWriter(new_file, fieldnames=fieldnames, delimiter=',')
                csv_writer.writeheader()
                for line in csv_reader:
                    if (line.get('Job Location', None) == district) and (((line.get('Item ID', None) == '55')
                            or (line.get('Item ID', None) == '56') or (line.get('Item ID', None) == '57') or
                            (line.get('Item ID', None) == '83') or (line.get('Item ID', None) == '2') or
                            (line.get('Item ID', None) == '32'))):
                        # print(line)
                        csv_writer.writerow(line)
            csv_file.seek(0)        # Reset to beginning of dictionary
    # Use pandas to create xlsx files
    needs_crane = pd.DataFrame
    needs_pc = pd.DataFrame
    needs_fork = pd.DataFrame
    needs_aerial = pd.DataFrame
    for name in f_name_list:
        print(name)
        needs = pd.read_csv(name)
        needs['Days Remaining'] = needs['Days Remaining'].apply(str)
        needs['Days Remaining'] = needs['Days Remaining'].str.replace(',', '')
        needs['Days Remaining'] = pd.to_numeric(needs['Days Remaining'])
        needs.columns = [col.replace(" ", "_") for col in needs]
        name = name.replace('.csv', '.xlsx')
        writer = pd.ExcelWriter(name, engine='xlsxwriter')
        needs.to_excel(writer, index=False)
        writer.save()
    print("File completed!")
    result_text.set("Results: File completed!")


# *************************** Create the Window Objects **************************


root = tk.Tk()  # ************ Main (root) Window
root.option_add('*tearOff', False)
root.title('Parse Training Needs Report')
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
parse_btn = ttk.Button(mainframe, text="Parse File", command=parse_file)
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
