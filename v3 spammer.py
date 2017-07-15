from Tkinter import *
import smtplib
import getpass
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from ScrolledText import *
import tkMessageBox
import urllib2

def quit():
    if tkMessageBox.askyesno("Quit", "Are you sure you want to quit?"):
        root.destroy()

root = Tk(className=" Email")
root.focus()
root.geometry("{}x{}".format(400, 250))
menu = Menu(root)
root.config(menu=menu)
filemenu = Menu(menu)
menu.add_cascade(label="File", menu=filemenu)
filemenu.add_command(label="Quit...", command=quit)

#root labels
useremaillabel = Label(root, text="Your email: ")
useremaillabel.grid(row=0, column=0)
subjectlabel = Label(root, text="Email subject: ")
subjectlabel.grid(row=1, column=0)
recipientemaillabel = Label(root, text="Enter recipient: ")
recipientemaillabel.grid(row=2, column=0)
msglabel = Label(root, text="Enter your message: ")
msglabel.grid(row=3, column=0)
numofemails = Label(root, text="Enter number of emails: ")
numofemails.grid(row=4, column=0)

#root entry boxes
useremailentry = Entry(root, bd=1)
useremailentry.grid(row=0, column=1)
subjectentry = Entry(root, bd=1)
subjectentry.grid(row=1, column=1)
recipiententry = Entry(root, bd=1)
recipiententry.grid(row=2, column=1)
msgentry = ScrolledText(root, width=23, height=4, bd=1)
msgentry.grid(row=3, column=1)
numofemailsentry = Entry(root, bd=1)
numofemailsentry.grid(row=4, column=1)

def gmail():     
    fromaddr = useremailentry.get()
    toaddr = recipiententry.get()
    subject = subjectentry.get()
    body = msgentry.get("1.0", 'end-1c')
    frompwd = getpass.getpass("Enter email password: ")
    gnum = numofemailsentry.get()
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    i = 0
    while i < int(gnum):
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.ehlo()
        server.login(fromaddr, frompwd)
        text = msg.as_string()
        server.sendmail(fromaddr, toaddr, text)
        server.quit()
        i += 1
        if i == int(gnum):
            print "Done, " + str(gnum) + " email(s) sent to " + str(toaddr)

def office():
    fromaddr = useremailentry.get()
    toaddr = recipiententry.get()
    subject = subjectentry.get()
    body = msgentry.get("1.0", 'end-1c')
    frompwd = getpass.getpass("Enter email password: ")
    gnum = numofemailsentry.get()
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    i = 0
    while i < int(gnum):
        server = smtplib.SMTP("smtp.office365.com", 587)
        server.starttls()
        server.ehlo()
        server.login(fromaddr, frompwd)
        text = msg.as_string()
        server.sendmail(fromaddr, toaddr, text)
        server.quit()
        i += 1
        if i == int(gnum):
            print "Done, " + str(gnum) + " email(s) sent to " + str(toaddr)

def outlook():
    fromaddr = useremailentry.get()
    toaddr = recipiententry.get()
    subject = subjectentry.get()
    body = msgentry.get("1.0", 'end-1c')
    frompwd = getpass.getpass("Enter email password: ")
    gnum = numofemailsentry.get()
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    i = 0
    while i < int(gnum):
        server = smtplib.SMTP("smtp-mail.outlook.com", 587)
        server.starttls()
        server.ehlo()
        server.login(fromaddr, frompwd)
        text = msg.as_string()
        server.sendmail(fromaddr, toaddr, text)
        server.quit()
        i += 1
        if i == int(gnum):
            print "Done, " + str(gnum) + " email(s) sent to " + str(toaddr)

def hotmail():
    fromaddr = useremailentry.get()
    toaddr = recipiententry.get()
    subject = subjectentry.get()
    body = msgentry.get("1.0", 'end-1c')
    frompwd = getpass.getpass("Enter email password: ")
    gnum = numofemailsentry.get()
    msg = MIMEMultipart()
    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    i = 0
    while i < int(gnum):
        server = smtplib.SMTP("smtp.live.com", 25)
        server.starttls()
        server.ehlo()
        server.login(fromaddr, frompwd)
        text = msg.as_string()
        server.sendmail(fromaddr, toaddr, text)
        server.quit()
        i += 1
        if i == int(gnum):
            print "Done, " + str(gnum) + " email(s) sent to " + str(toaddr)

#root buttons
gmailbutton = Button(root, text="Gmail", command=gmail)
gmailbutton.grid(row=6, column=0)
officebutton = Button(root, text="Office 365", command=office)
officebutton.grid(row=6, column=1)
outlookbutton = Button(root, text="Outlook", command=outlook)
outlookbutton.grid(row=7, column=0)
hotmailbutton = Button(root, text="Hotmail", command=hotmail)
hotmailbutton.grid(row=7, column=1)


root.mainloop()
