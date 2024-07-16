import os
import smtplib
from email.message import EmailMessage
from docx2pdf import convert
from docxtpl import DocxTemplate
import datetime
import tkinter
import random
from tkinter import ttk
from tkinter import messagebox
from tkcalendar import DateEntry
from PIL import ImageGrab 

# Define the directory where invoices will be saved
invoice_dir = "word"
os.makedirs(invoice_dir, exist_ok=True)
pdf_dir = "pdfss"

os.makedirs(pdf_dir, exist_ok=True)

def clear_item():
    desc_entry.delete(0, tkinter.END)
    qty_entey.delete(0, tkinter.END)
    price_entry.delete(0, tkinter.END)
    order_id_entry.delete(0, tkinter.END)


def generate_random_invoice_number():
    random_number = random.randint(1, 100000000)
    invoice_number_entry.delete(0, tkinter.END)
    invoice_number_entry.insert(0, str(random_number))



invoice_list = []
def add_item():
    order = order_id_entry.get()
    desc = desc_entry.get()
    qty = int(qty_entey.get())
    price = float(price_entry.get())
    total = qty*price
    invoice_item =[order,desc,qty,price,total]

    tree.insert('', 0, values=invoice_item)
    clear_item()
    invoice_list.append(invoice_item)


def new_invoice():
    name_entry.delete(0,tkinter.END)
    email_entry.delete(0, tkinter.END)
    phone_entry.delete(0, tkinter.END)
    address_entry.delete(0, tkinter.END)

    clear_item()
    tree.delete(*tree.get_children())

    invoice_list.clear()


def generate_invoice():
    doc = DocxTemplate("yyyy.docx")
    name = name_entry.get()
    phone = phone_entry.get()
    email = email_entry.get()
    address = address_entry.get()
    invoice_no = invoice_number_entry.get()
    date = call_Date.get()
    #desc = desc_entry.get()
    #qty = qty_entey.get()
    #price = price_entry.get()
    order = order_id_entry.get()
    subtotal = sum(item[4] for item in invoice_list)
    tax = 0
    #total = subtotal*(1-tax)
    total = subtotal




    doc.render({
        'name': name,
        'phone': phone,
        'email': email,
        'address': address,
        'invoice': invoice_no,
        'date' : date,
        'order': order,
        'invoice_list': invoice_list,
        'subtotal': subtotal,
        'tax': str(subtotal*100)+"%",
        'total': total
    })

    
    file_name = os.path.join(invoice_dir, f"{invoice_no}.docx")
    doc.save(file_name)
    messagebox.showinfo("Success", f"Invoice saved as docx successfully!")

    #doc_name =name +datetime.datetime.now().strftime("%Y-%m-%d-%H%M%S")+".docx"
    #doc_name = invoice_no +".pdf"
    #doc.save(doc_name)

def save_as_pdf():
    docx_file = os.path.join(invoice_dir, f"{invoice_number_entry.get()}.docx")
    pdf_file = os.path.join(invoice_dir, f"{invoice_number_entry.get()}.pdf")

    if not os.path.exists(docx_file):
        messagebox.showwarning("File Error", "Invoice file does not exist. Generate the invoice first.")
        return

    try:
        # Convert docx to pdf
        convert(docx_file, pdf_file)
        messagebox.showinfo("Success", f"Invoice saved as PDF successfully!")
    except Exception as e:
        messagebox.showerror("Conversion Error", f"An error occurred while converting to PDF: {e}")

def save_as_png():
    # Capture the window and save as PNG
    x = window.winfo_rootx()
    y = window.winfo_rooty()
    w = window.winfo_width()
    h = window.winfo_height()
    ImageGrab.grab(bbox=(x, y, x + w, y + h)).save(os.path.join(invoice_dir, f"{invoice_number_entry.get()}.png"))
    messagebox.showinfo("Success", "Invoice saved as PNG successfully!")

def send_email():
    email_address = "your_email@exmple.com"
    email_password = "your_password"
    recipient_email = email_entry.get()

    if not recipient_email:
        messagebox.showwarning("Input Error", "please enter the recipient's email address")
        return
    
    subject = "Invoice"
    body = "Please find attached your invoice."
    file_name = os.path.join(invoice_dir, f"invoice_{invoice_number_entry.get()}.docx")

    if not os.path.exists(file_name):
        messagebox.showwarning("File Error", "Invoice file does not exist. Generate the invoice first.")
        return
    msg = EmailMessage()
    msg['From'] = email_address
    msg['To'] = recipient_email
    msg['Subject'] = subject
    msg.set_content(body)

    with open(file_name, 'rb') as f:
        file_data = f.read()
        file_name = f.name
        msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(email_address, email_password)
        smtp.send_message(msg)

    messagebox.showinfo("Success", "Email sent successfully!")

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
generate_number_button = tkinter.Button(frame, text="Generate Number", command=generate_random_invoice_number)
generate_number_button.grid(row=4, column=0, padx=20,pady=10)
order_id_label = tkinter.Label(frame, text="Order ID")
order_id_label.grid(row=4, column=1)
order_id_entry = tkinter.Entry(frame)
order_id_entry.grid(row=5, column=1)

#create a date entry
call_Date = DateEntry(frame, width=12, background="white", foreground="", borderwidt=2)
call_Date.grid(row=5, column=2)


#get_data_button = tkinter.Button(frame, text="Get selected Date")
#get_data_button.grid(row=5, column=3)


desc_entry = tkinter.Entry(frame)
desc_entry.grid(row=3, column=1)
qty_entey = tkinter.Spinbox(frame, from_=0, to= 1000)
qty_entey.grid(row=3, column=2)
price_entry = tkinter.Spinbox(frame, from_=0.0, to=100000000, increment=2)
price_entry.grid(row=3, column=3)


add_item_button = tkinter.Button(frame, text="Add Item", command=add_item )
add_item_button.grid(row=5, column=3,padx=20,pady=10)

columns = ('order','desc', 'qty', 'price', 'total')
tree = ttk.Treeview(frame, columns=columns, show="headings")
tree.heading('order', text='Order ID')
tree.heading('desc', text='Description')
tree.heading('qty', text='Qty')
tree.heading('price', text='Price')
tree.heading('total', text='Total')
tree.grid(row=6,column=0, columnspan=4, pady=20, padx=10)

save_invoice_button = tkinter.Button(frame, text="Generate Invoice", command=generate_invoice) 
save_invoice_button.grid(row=7, column=0)

save_pdf_button = tkinter.Button(frame, text="Pdf Save", command=save_as_pdf)
save_pdf_button.grid(row=7, column=1)

save_png_button = tkinter.Button(frame, text= "Save to png", command=save_as_png)
save_png_button.grid(row=7, column=2)

send_gmail_button = tkinter.Button(frame, text="Email Invoice", command=send_email)
send_gmail_button.grid(row=7, column=3)

new_invoice_button = tkinter.Button(frame, text="New Invoice", command=new_invoice)
new_invoice_button.grid(row=7, column=4, padx=0, pady=5)


window.mainloop()