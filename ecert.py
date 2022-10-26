from tkinter import *
from tkinter import filedialog, StringVar
from tkinter.constants import X
import pandas as pd
import smtplib
import string
import time
import tkinter as tk
import pandas as pd
import sys
import webbrowser
import os
import shutil
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from PIL import Image, ImageDraw, ImageFont, ImageTk
from pandas import ExcelWriter
from pandas import ExcelFile
import re
#p_wE_d = "Ihealthecert"
#fromaddr = "ikramhealth.ecert@gmail.com"

root = tk.Tk()
root.title("MEduCare eCert")

canvas1 = tk.Canvas(root, width=500, height=500, bg='lightsteelblue')
canvas1.pack()
myvar = tk.StringVar()

def callback(url):
    webbrowser.open_new(url)


def openSettingWindow():

    newWindow = Toplevel(root)
    newWindow.title("Settings")
    newWindow.geometry("300x200")

    f = open("txt/emailsetting.txt", "r")
    emailsetting = f.read().split(',')

    def save_setting():
        s = open("txt/emailsetting.txt", "w")
        email_text = email_tw.get(1.0, tk.END)
        pwd_text = pwd_tw.get(1.0, tk.END)
        settingnew = str(email_text) + ',' + str(pwd_text)
        #pwd_text = pwd_tw.get(1.0, tk.END)

        s.write(settingnew)
        newWindow.destroy()

    l1 = Label(newWindow, text="Edit email")
    l1.pack()

    email_tw = Text(newWindow, width=30, height=1)
    email_tw.insert(tk.END, emailsetting[0])
    email_tw.pack(pady=10)

    l2 = Label(newWindow, text="Edit password")
    l2.pack()

    pwd_tw = Text(newWindow, width=30, height=1)
    pwd_tw.insert(tk.END, emailsetting[1])
    pwd_tw.pack()

    saveButton = Button(newWindow, text="SAVE", command=save_setting, width=15)
    saveButton.pack(pady=5)

    cancelButton = Button(newWindow, text="CANCEL",
                          command=newWindow.destroy, width=15)
    cancelButton.pack()


def getExcel():
    try:
        global df
        import_file_path = filedialog.askopenfilename()
        df = pd.read_excel(import_file_path)
    except:
        pass


def autoCert():
    f = open("txt/emailsetting.txt", "r")
    emailsetting = f.read().split(',')
    p_wE_d = emailsetting[1]
    fromaddr = emailsetting[0]
    cert_path = "certs"
    isFile = os.path.isdir(cert_path)
    if isFile == False:
        os.makedirs(cert_path)
    else:
        shutil.rmtree(cert_path)
        os.makedirs(cert_path)
    
    try:

        for i in df.index:
            image = Image.open('template.jpg')
            draw = ImageDraw.Draw(image)
            font = ImageFont.truetype('Cinzel-Bold.otf', size=60)
            color = 'rgb(0, 0, 0)'
            name = df['Name'][i]
            name = name.upper()
            name = name.replace('/', '')
            print(i+1, name)
            w, h = draw.textsize(name, font=font)
            draw.text(((image.width - w)/2, 530), name, fill=color, font=font)
            imageName = "certs/"+name+".pdf"
            image.save(imageName)
            time.sleep(1)

            toaddr = df['Email'][i]
            msg = MIMEMultipart()
            msg['From'] = fromaddr

            # storing the receivers email address
            msg['To'] = toaddr

            # storing the subject
            f = open("txt/subject.txt", "r")
            subject_text = f.read()
            msg['Subject'] = subject_text

            # string to store the body of the mail
            f = open("txt/msg.txt", "r")
            body = f.read()

            # attach the body with the msg instance
            msg.attach(MIMEText(body, 'plain'))

            # open the file to be sent
            filename = name + ".pdf"
            attachment = open(imageName, "rb")
            p = MIMEBase('application', 'octet-stream')

            # To change the payload into encoded form
            p.set_payload((attachment).read())

            # encode into base64
            encoders.encode_base64(p)

            #p.add_header('Content-Disposition',"attachment; filename= %s" % filename)
            p.add_header('Content-Disposition', 'attachment',filename='%s' % filename)
            # attach the instance 'p' to instance 'msg'
            msg.attach(p)

            # creates SMTP session
            s = smtplib.SMTP('smtp.gmail.com', 587)

            # start TLS for security
            s.starttls()

            # Authentication
            s.login(fromaddr, p_wE_d)

            # Converts the Multipart msg into a string
            text = msg.as_string()

            # sending the mail
            s.sendmail(fromaddr, toaddr, text)
            time.sleep(1)

            # terminating the session

        s.quit()
        print("PROCESS COMPLETED")
        input()
    #except smtplib.SMTPAuthenticationError:
        #print(e)
        #print("Less secure apps not enabled. Please go to https://myaccount.google.com/lesssecureapps and ENABLE Less Secure App under Ihealthecert@gmail.com account")

    except FileNotFoundError:
        #print(e)
        print("Please create 'certs' folder")

    except Exception as e:
        print(e)
        print("ERROR,NOT ABLE TO PROCEED")
        pass


class DynamicImage(tk.Label):
    def __init__(self, master=None, image_path="", *args, **kwargs):
        super().__init__(master, *args, **kwargs)
        self.width = master.winfo_screenwidth()//2
        self.height = master.winfo_screenheight()//2
        self.img = Image.open(image_path)
        self.p_img = None
        self.bind("<Configure>", self.resizing)

    def resizing(self, event=None):
        new_window = Toplevel(root)
        w, h = self.img.width, self.img.height
        if w > h:
            delta = self.width/w
            new_width, new_height = self.width, int(h*delta)
        else:
            delta = self.height/h
            new_width, new_height = int(w*delta), self.height
        self.p_img = ImageTk.PhotoImage(
            self.img.resize((new_width, new_height)))
        image = self.p_img
        panel = Label(new_window, image=image,
                      width=new_width, height=new_height)
        panel.image = self.p_img
        panel.pack()


def createCert():
    image = Image.open('template.jpg')
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype('Cinzel-Bold.otf', size=60)
    color = 'rgb(0, 0, 0)'
    w, h = draw.textsize(myvar.get(), font=font)
    draw.text(((image.width - w)/2, 530), myvar.get(), fill=color, font=font)
    imageName1 = myvar.get() + ".png"
    image.save(imageName1)


def changeTemplate():
    image = Image.open("raw.jpg")
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype('Cinzel-Bold.otf', size=60)
    color = 'rgb(0, 0, 0)'
    w, h = draw.textsize(coursevar.get(), font=font)
    x, h = draw.textsize(datevar.get(), font=font)
    draw.text(((image.width - w)/2, 735),
              coursevar.get(), fill=color, font=font)
    draw.text(((image.width - x)/2, 945), datevar.get(), fill=color, font=font)
    imageName2 = "template.jpg"
    image.save(imageName2)


def openNewWindow():

    newWindow = Toplevel(root)
    newWindow.title("Edit email text")
    newWindow.geometry("600x500")

    f = open("txt/subject.txt", "r")
    subject_t = f.read()

    g = open("txt/msg.txt", "r")
    msg_t = g.read()

    def save_subject():
        s = open("txt/subject.txt", "w")
        sub_text = subject_tw.get(1.0, tk.END)
        s.write(sub_text)
        w = open("txt/msg.txt", "w")
        msg_text = msg_tw.get(1.0, tk.END)
        w.write(msg_text)
        newWindow.destroy()

    l1 = Label(newWindow, text="Edit subject")
    l1.pack()

    subject_tw = Text(newWindow, width=75, height=1)
    subject_tw.insert(tk.END, subject_t)
    subject_tw.pack(pady=10)

    l2 = Label(newWindow, text="Edit message")
    l2.pack()

    msg_tw = Text(newWindow, width=75, height=20)
    msg_tw.insert(tk.END, msg_t)
    msg_tw.pack()

    saveButton = Button(newWindow, text="SAVE", command=save_subject, width=15)
    saveButton.pack(pady=5)

    cancelButton = Button(newWindow, text="CANCEL",
                          command=newWindow.destroy, width=15)
    cancelButton.pack()


def excelName():
    try:
        if "Name" in df:
            excel_var.set('Excel file Loaded')
        else:
            excel_var.set('Excel file not found. Please select excel file')
    except NameError:
        excel_var.set('Excel file not found. Please select excel file')
        print('Please select excel file')


def test1():
    if "Name" in df:
        name_var.set('Name column check')
    else:
        name_var.set('Name column not present')


def test2():
    if "Email" in df:
        for i in df.index:
            toaddr = df['Email'][i]
            pat = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
            if re.match(pat,toaddr):
                pass
            else:
                email_var.set('Invalid email detected. Please edit email column')
                break
            email_var.set('Email column check')

    else:
        email_var.set('Email column not present')


def autocert_f():
    # This will remove the widget from toplevel
    canvas1.itemconfig(etext, state='hidden')
    canvas1.itemconfig(run_onecert, state='hidden')
    canvas1.itemconfig(name_onecert, state='hidden')
    canvas1.itemconfig(excel_l, state='normal')
    canvas1.itemconfig(name_l, state='normal')
    canvas1.itemconfig(email_l, state='normal')
    canvas1.itemconfig(browse_B, state='normal')
    canvas1.itemconfig(run_autocert, state='normal')
    canvas1.itemconfig(editBtn, state='normal')
    canvas1.itemconfig(course_l, state='hidden')
    canvas1.itemconfig(course_e, state='hidden')
    canvas1.itemconfig(date_l, state='hidden')
    canvas1.itemconfig(date_e, state='hidden')
    canvas1.itemconfig(run_temp, state='hidden')
    canvas1.itemconfig(temp_n, state='normal')
    #canvas1.itemconfig(lesssecure_l, state='normal')


def onecert_f():
    # This will remove the widget from toplevel
    canvas1.itemconfig(etext, state='normal')
    canvas1.itemconfig(run_onecert, state='normal')
    canvas1.itemconfig(name_onecert, state='normal')
    canvas1.itemconfig(excel_l, state='hidden')
    canvas1.itemconfig(name_l, state='hidden')
    canvas1.itemconfig(email_l, state='hidden')
    canvas1.itemconfig(browse_B, state='hidden')
    canvas1.itemconfig(editBtn, state='hidden')
    canvas1.itemconfig(run_autocert, state='hidden')
    canvas1.itemconfig(course_l, state='hidden')
    canvas1.itemconfig(course_e, state='hidden')
    canvas1.itemconfig(date_l, state='hidden')
    canvas1.itemconfig(date_e, state='hidden')
    canvas1.itemconfig(run_temp, state='hidden')
    canvas1.itemconfig(temp_n, state='hidden')
    #canvas1.itemconfig(lesssecure_l, state='hidden')


def temp_f():
    # This will remove the widget from toplevel
    canvas1.itemconfig(etext, state='hidden')
    canvas1.itemconfig(run_onecert, state='hidden')
    canvas1.itemconfig(name_onecert, state='hidden')
    canvas1.itemconfig(excel_l, state='hidden')
    canvas1.itemconfig(name_l, state='hidden')
    canvas1.itemconfig(email_l, state='hidden')
    canvas1.itemconfig(browse_B, state='hidden')
    canvas1.itemconfig(editBtn, state='hidden')
    canvas1.itemconfig(run_autocert, state='hidden')
    canvas1.itemconfig(course_l, state='normal')
    canvas1.itemconfig(course_e, state='normal')
    canvas1.itemconfig(date_l, state='normal')
    canvas1.itemconfig(date_e, state='normal')
    canvas1.itemconfig(run_temp, state='normal')
    canvas1.itemconfig(temp_n, state='hidden')
    #canvas1.itemconfig(lesssecure_l, state='hidden')

excel_var = StringVar()
name_var = StringVar()
email_var = StringVar()
coursevar = tk.StringVar()
datevar = tk.StringVar()

#########TITLE#############
label1 = tk.Label(text='MEduCare eCert', bg='lightsteelblue',
                  fg='black', font=('helvetica', 40, 'bold'))
canvas1.create_window(250, 90, window=label1)

created_by = tk.Label(text='version 1.3 created by sufian safaai',
                      bg='lightsteelblue', fg='black', font=('helvetica', 8, 'bold'))
canvas1.create_window(250, 130, window=created_by)

##########MENU BUTTONS#############
autocert_nav = tk.Button(text='Batch Cert', command=autocert_f,
                         bg='grey', fg='white', font=('helvetica', 8, 'bold'), width=12)
autocert_n = canvas1.create_window(125, 160, window=autocert_nav)

onecert_nav = tk.Button(text='Test Cert', command=onecert_f,
                        bg='grey', fg='white', font=('helvetica', 8, 'bold'), width=12)
onecert_n = canvas1.create_window(250, 160, window=onecert_nav)

setting_nav = tk.Button(text='Settings', command=openSettingWindow,
                        bg='grey', fg='white', font=('helvetica', 8, 'bold'), width=12)
onecert_n = canvas1.create_window(375, 160, window=setting_nav)


#######AUTOGENERATE CERT###########

temp_nav = tk.Button(text='Step 1:    Change Template', command=temp_f,
                     bg='green', fg='white', font=('helvetica', 12, 'bold'), width=25)
temp_n = canvas1.create_window(250, 210, window=temp_nav)

edittextButton = tk.Button(text='Step 2: Edit email message', command=openNewWindow,
                           bg='green', fg='white', font=('helvetica', 12, 'bold'), width=25)
editBtn = canvas1.create_window(250, 250, window=edittextButton)

browseButton_Excel = tk.Button(text='Step 3: Import Excel File', command=lambda: [getExcel(), excelName(), test1(), test2()], bg='green', fg='white', font=('helvetica', 12, 'bold'), width=25)
browse_B = canvas1.create_window(250, 290, window=browseButton_Excel)

excel_label = tk.Label(textvariable=excel_var, bg='lightsteelblue',
                       fg='black', font=('helvetica', 10, 'bold'))
excel_l = canvas1.create_window(250, 320, window=excel_label)

name_label = tk.Label(textvariable=name_var, bg='lightsteelblue',
                      fg='black', font=('helvetica', 10, 'bold'))
name_l = canvas1.create_window(250, 340, window=name_label)

email_label = tk.Label(textvariable=email_var, bg='lightsteelblue',
                       fg='black', font=('helvetica', 10, 'bold'))
email_l = canvas1.create_window(250, 360, window=email_label)

#lesssecure_label = tk.Label(text='Click here to ENABLE LESS SECURE APPS under the gmail account', bg='lightsteelblue',
#                       fg='blue', font=('helvetica', 8, 'bold'))
#lesssecure_label.bind("<Button-1>", lambda e: callback("https://myaccount.google.com/lesssecureapps"))
#lesssecure_l = canvas1.create_window(250, 380, window=lesssecure_label)

runButton = tk.Button(text='Final step:    Run AutoCert', command=autoCert,
                      bg='maroon', fg='white', font=('helvetica', 12, 'bold'), width=25)
run_autocert = canvas1.create_window(250, 430, window=runButton)


############ONE CERT#############

name_label = tk.Label(text='Enter name', bg='lightsteelblue',
                      fg='black', font=('helvetica', 10, 'bold'))
name_onecert = canvas1.create_window(
    250, 210, window=name_label, state='hidden')

entry_text = tk.Entry(textvariable=myvar, font=(
    'Cinzel-Bold.otf', 20, 'bold'), width=30)
etext = canvas1.create_window(250, 240, window=entry_text, state='hidden')

runButton = tk.Button(text='Generate A Cert', command=lambda: [createCert(), DynamicImage(root, image_path=myvar.get(
) + ".png").pack(fill="both", expand=True)], bg='green', fg='white', font=('helvetica', 12, 'bold'), width=15)
run_onecert = canvas1.create_window(250, 330, window=runButton, state='hidden')

##########CHANGE TEMPLATE###########

course_label = tk.Label(text='Enter course name', bg='lightsteelblue',
                        fg='black', font=('helvetica', 10, 'bold'))
course_l = canvas1.create_window(250, 190, window=course_label, state='hidden')

course_entry = tk.Entry(textvariable=coursevar, font=(
    'Cinzel-Bold.otf', 20, 'bold'), width=30)
course_e = canvas1.create_window(250, 220, window=course_entry, state='hidden')


date_label = tk.Label(text='Enter course date', bg='lightsteelblue',
                      fg='black', font=('helvetica', 10, 'bold'))
date_l = canvas1.create_window(250, 260, window=date_label, state='hidden')

date_entry = tk.Entry(textvariable=datevar, font=(
    'Cinzel-Bold.otf', 20, 'bold'), width=15)
date_e = canvas1.create_window(250, 290, window=date_entry, state='hidden')

run_template = tk.Button(text='Change template', command=lambda: [changeTemplate(), autocert_f(), DynamicImage(
    root, image_path="template.jpg").pack(fill="both", expand=True)], bg='green', fg='white', font=('helvetica', 12, 'bold'), width=15)
run_temp = canvas1.create_window(250, 350, window=run_template, state='hidden')


root.mainloop()