import tkinter as tk
from tkinter import ttk,messagebox
import openpyxl
from openpyxl import load_workbook

window = tk.Tk()
window.title("Simple Registration form")
window.geometry("300x500")

file_path=r"Book1.xlsx"
a=openpyxl.load_workbook(file_path)
b=a["Sheet1"]


def on_click():
    name=name_entry.get()
    email=email_entry.get()
    age=age_entry.get()
    gender=g.get()
    id=id_entry.get()
    ph=ph_entry.get()
    year=year_entry.get()
    branch=branch_entry.get()
    agree=val.get()

    if name and email and id and ph and year and branch and agree and gender and age:
        if "@" not in email:
            messagebox.showerror("error", "email incorrect")
        else :
            messagebox.showinfo("submitted","name:{}\nemail:{}\nGender:{}\nage:{}\nid:{}\nph:{}\nyear:{}\nbranch:{}\nagree:{}".format(name,email,gender,age,id,ph,year,branch,agree))
            b.append([name,email,gender,age,id,ph,year,branch,agree])
            a.save(file_path)

    else:
         messagebox.showwarning("warning","fill all the entries")


name_label=tk.Label(text="Name:")
name_label.grid(row=0,column=0,padx=10,pady=10)
name_entry=tk.Entry()
name_entry.grid(row=0,column=1)

email_label=tk.Label(text="Email:")
email_label.grid(row=1,column=0,padx=10,pady=10)
email_entry=tk.Entry()
email_entry.grid(row=1,column=1)

g=tk.StringVar()
gender_label=tk.Label(text="Gender:")
gender_label.grid(row=2,column=0,padx=10,pady=10)
gender_entry_m=ttk.Radiobutton(text="Male",value="male",variable=g)
gender_entry_m.grid(row=2,column=1)
gender_entry_f=ttk.Radiobutton(text="Female",value="female",variable=g)
gender_entry_f.grid(row=3,column=1)

age_label=tk.Label(text="Age:")
age_label.grid(row=4,column=0,padx=10,pady=10)
age_entry=tk.Spinbox(from_=15,to=25)
age_entry.grid(row=4,column=1)


ph_label=tk.Label(text="Ph Number:")
ph_label.grid(row=5,column=0,padx=10,pady=10)
ph_entry=tk.Entry()
ph_entry.grid(row=5,column=1)

id_label=tk.Label(text="ID Number:")
id_label.grid(row=6,column=0,padx=10,pady=10)
id_entry=tk.Entry()
id_entry.grid(row=6,column=1)

year_label=tk.Label(text="year:")
year_label.grid(row=7,column=0,padx=10,pady=10)
year_entry= ttk.Combobox(values=["1st","2nd","3rd","4th"])
year_entry.grid(row=7,column=1)

branch_label=tk.Label(text="Branch:")
branch_label.grid(row=8,column=0,padx=10,pady=10)
branch_entry= ttk.Combobox(values=["ECE","CSE","CIVIL","MECH","EEE"])
branch_entry.grid(row=8,column=1)


val=tk.IntVar()
agree_entry =tk.Checkbutton(text="agree TC",variable=val)
agree_entry.grid(row=9,column=1,padx=10,pady=10)

submit = tk.Button(text="Submit",command=on_click)
submit.grid(row=10,column=1,padx=10,pady=10)

window.mainloop()
