import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import sys
import pandas as pd

import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

from PIL import Image, ImageTk

class PrintLogger: 
    def __init__(self, textbox): 
        self.textbox = textbox 
 
    def write(self, text): 
        self.textbox.insert(tk.END, text) 

def browse_attachment_file():
    file_path = filedialog.askopenfilename(filetypes=[("Any file", "*")])
    if file_path:
        entry_attachment_path.delete(0, tk.END)
        entry_attachment_path.insert(0, file_path)

def browse_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        entry_excel_path.delete(0, tk.END)
        entry_excel_path.insert(0, file_path)

def send_email(sender_email, sender_name, app_password, receiver_email, subject, body, attachment_path):
    try:
        # instance of MIMEMultipart
        msg = MIMEMultipart()
        msg['To'] = receiver_email
        # storing the senders email address  
        msg['From'] = sender_name
        # storing the subject 
        msg['Subject'] = subject
        # attach the body with the msg instance
        msg.attach(MIMEText(body, 'plain')) 
        if attachment_path != '':
            # open the file to be sent 
            attachment = open(attachment_path, "rb") 
            # instance of MIMEBase and named as p
            p = MIMEBase('application', 'octet-stream') 
            # To change the payload into encoded form
            p.set_payload((attachment).read())
            # encode into base64
            encoders.encode_base64(p)
            p.add_header('Content-Disposition', "attachment; filename= %s" % os.path.basename(attachment_path))
            # attach the instance 'p' to instance 'msg'
            msg.attach(p)
        s = smtplib.SMTP('smtp.gmail.com', 587) # creates SMTP session
        s.starttls() # start TLS for security
        s.login(sender_email, app_password) # authenticate
        s.sendmail(sender_email, receiver_email, msg.as_string())
        s.quit()
    except Exception as e:
        print(f"Failed to send email: {str(e)}")
        messagebox.showerror("Failed to send email to {}".format(receiver_email), "Failed to send email to {}\n".format(receiver_email) + str(e))

def send_bulk_emails():
    sender_email = entry_sender_email.get()
    sender_name = entry_sender_name.get()
    sender_password = entry_sender_password.get()
    subject = entry_subject.get()
    excel_file_path = entry_excel_path.get()
    attachment_file_path = entry_attachment_path.get()

    receivers = pd.read_excel(excel_file_path)
    total_emails = receivers.shape[0]
    print('Sending {} emails'.format(total_emails))
    for idx, receiver_row in receivers.iterrows():
        print('Sending email {count}/{total}: {receiver}'.format(count = idx + 1, total = total_emails, receiver = receiver_row['Email']))
        body = text_body.get("1.0", tk.END).format(receiver_row['First'])
        send_email(sender_email, sender_name, sender_password, receiver_row['Email'], subject, body, attachment_file_path)
    print('Done! Now go and enjoy San Diego ;)')


app = tk.Tk()
app.iconbitmap("octopus_icon.ico")
app.title("Anna's Bulk Email Sender")

label_sender_email = tk.Label(app, text="Sender Email:")
label_sender_email.grid(row=0, column=0)
entry_sender_email = tk.Entry(app)
entry_sender_email.grid(row=0, column=1)

label_sender_name = tk.Label(app, text="Sender Name (to show in the \'From' section):")
label_sender_name.grid(row=1, column=0)
entry_sender_name = tk.Entry(app)
entry_sender_name.grid(row=1, column=1)

label_sender_password = tk.Label(app, text="Sender Password:")
label_sender_password.grid(row=2, column=0)
entry_sender_password = tk.Entry(app, show="*")
entry_sender_password.grid(row=2, column=1)

label_subject = tk.Label(app, text="Subject:")
label_subject.grid(row=3, column=0)
entry_subject = tk.Entry(app)
entry_subject.grid(row=3, column=1)

label_body = tk.Label(app, text="Body:")
label_body.grid(row=4, column=0)
text_body = tk.Text(app, height=10, width=30)
text_body.grid(row=4, column=1)

label_attachment_path = tk.Label(app, text="Attach file:")
label_attachment_path.grid(row=5, column=0)
entry_attachment_path = tk.Entry(app)
entry_attachment_path.grid(row=5, column=1)
btn_browse_attachment = tk.Button(app, text="Browse", command=browse_attachment_file)
btn_browse_attachment.grid(row=5, column=2)

label_excel_path = tk.Label(app, text="Excel file with email list:")
label_excel_path.grid(row=6, column=0)
entry_excel_path = tk.Entry(app)
entry_excel_path.grid(row=6, column=1)
btn_browse_excel = tk.Button(app, text="Browse", command=browse_excel_file)
btn_browse_excel.grid(row=6, column=2)

btn_send_emails = tk.Button(app, text="Send Emails", command=send_bulk_emails)
btn_send_emails.grid(row=7, column=0, columnspan=3)

textbox = tk.Text(app) 
textbox.grid(column=0, row=8, columnspan=3)
 
printlogger = PrintLogger(textbox) 
sys.stdout = printlogger

app.mainloop()
