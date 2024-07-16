import tkinter
from tkinter import ttk
from docxtpl import DocxTemplate
import datetime
from tkinter import messagebox
import random


window = tkinter.Tk()
window.title("Invoice Generator from")
window.iconbitmap("ico.ico")

frame = tkinter.Frame(window)
frame.pack(padx=20, pady=10)

name_label = tkinter.Label(frame, text="Name")
name_label.grid(row=0, column=0)
phone_label = tkinter.Label(frame, text="Phone")
phone_label.grid(row=0, column=1)
email_label = tkinter.Label(frame, text="Email")
email_label.grid(row=0, column=2)
address_label = tkinter.Label(frame, text="Address")
address_label.grid(row=0, column=3)

name_entry = tkinter.Entry(frame)
name_entry.grid(row=1, column=0)
phone_entry = tkinter.Entry(frame)
phone_entry.grid(row=1, column=1)
email_entry = tkinter.Entry(frame)
email_entry.grid(row=1, column=2)
address_entry = tkinter.Entry(frame)
address_entry.grid(row=1, column=3)

invoice_number_label = tkinter.Label(frame, text="Invoice Number")
invoice_number_label.grid(row=2, column=0)
desc_label = tkinter.Label(frame, text="Desc")
desc_label.grid(row=2, column=1)
qty_label = tkinter.Label(frame, text="Qty")
qty_label.grid(row=2, column=2)
price_label = tkinter.Label(frame, text="Price")
price_label.grid(row=2, column=3)

invoice_number_entry = tkinter.Entry(frame)
invoice_number_entry.grid(row=3, column=0)




window.mainloop()