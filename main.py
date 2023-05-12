import tkinter as tk
from tkinter import ttk
import openpyxl as xl

def toggle_mode():
    if mode_switch.instate(["selected"]):
        style.theme_use("forest-light")
    else:
        style.theme_use("forest-dark")
        

def load_data():
    wb = xl.load_workbook("Employers.xlsx")
    sheet = wb["Sheet1"]
    
    list_values = list(sheet.values)
    print(list_values)
    
    # for col_name in list_values[0]:
    #     print(col_name)
    #     treeview.heading(col_name, text=col_name)
    
    for value_tuple in list_values[1:]:
        treeview.insert('', tk.END, values=value_tuple)
        
def insert_row():
    
    name = name_entry.get()
    age = int (age_spinbox.get())
    subscription_status = status_combobox.get()
    employment_status = "Employed" if a.get() else "Unemployed"
    # Insert into the Excel sheet
    
    wb = xl.load_workbook("Employers.xlsx")
    sheet = wb["Sheet1"]
    row_values = [name, age, subscription_status, employment_status]
    sheet.append(row_values)
    wb.save("Employers.xlsx")
    
    # Insert into the treewiew
    
    treeview.insert("", tk.END, values=row_values)
    
    # Clear the values
    
    name_entry.delete(0, "end")
    name_entry.insert(0, "Name")
    age_spinbox.delete(0, "end")
    age_spinbox.insert(0, "Age")
    status_combobox.set(combo_list[0])
    checkbutton.state(["!selected"])
        

root = tk.Tk()
root.title("EXCEL APP")
root.iconbitmap("excel.ico")


style = ttk.Style(root)
root.tk.call("source", "forest-light.tcl")
root.tk.call("source", "forest-dark.tcl")
style.theme_use("forest-dark")

combo_list = ["Subscribed", "Not Subscribed", "Other"]

frame = ttk.Frame(root)
frame.pack()

widget_frame = ttk.LabelFrame(frame, text="Insert Box")
widget_frame.grid(row=0, column=0, padx=20, pady=10)

name_entry = ttk.Entry(widget_frame)
name_entry.insert(0, "Name")
name_entry.bind("<FocusIn>", lambda e: name_entry.delete("0", "end"))
name_entry.grid(row=0, column=0, padx=5, pady=(0, 5), sticky="ew")

age_spinbox = ttk.Spinbox(widget_frame, from_=18, to=100 )
age_spinbox.insert(0, "Age")
age_spinbox.grid(row=1, column=0,padx=5, pady=(0, 5), sticky="ew")

status_combobox = ttk.Combobox(widget_frame, values=combo_list)
status_combobox.current(0)
status_combobox.grid(row=2, column=0,padx=5, pady=(0, 5), sticky="ew")

a = tk.BooleanVar()
checkbutton = ttk.Checkbutton(widget_frame, text="Employed", variable=a)
checkbutton.grid(row=3, column=0,padx=5, pady=(0, 5), sticky="nsew")

insertbutton = ttk.Button(widget_frame, text="Insert", command=insert_row)
insertbutton.grid(row=4, column=0,padx=5, pady=(0, 5), sticky="nsew")

separator = ttk.Separator(widget_frame)
separator.grid(row=5, column=0, padx=(20,10), pady=10, sticky="ew")

mode_switch = ttk.Checkbutton(widget_frame, text="Mode", style="Switch", command=toggle_mode)
mode_switch.grid(row=6, column=0, padx=5, pady=10, sticky="nsew")

treeframe = ttk.Frame(frame)
treeframe.grid(row=0, column=1, pady=10)

treeScroll = ttk.Scrollbar(treeframe)
treeScroll.pack(side="right",fill="y")
cols = ("Name", "Age", "Subscription", "Employment")
treeview = ttk.Treeview(treeframe, show="headings", yscrollcommand=treeScroll.set, columns=cols, height=13)
treeview.column("Name", width=100)
treeview.column("Age", width=50)
treeview.column("Subscription", width=100)
treeview.column("Employment", width=100)
treeview.heading("Name", text="Name")
treeview.heading("Age", text="Age")
treeview.heading("Subscription", text="Subscription")
treeview.heading("Employment", text="Employment")
treeview.pack()
treeScroll.config(command=treeview.yview)

load_data()





root.mainloop()