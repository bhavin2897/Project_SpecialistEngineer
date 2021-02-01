import tkinter as tk
from tkinter import *
from tkinter import messagebox
from functools import partial
# from PIL import ImageTk, Image
import os

# Tkinter GUI
NUi = tk.Tk()
NUi.title("Niehoff Report Request")


# EmptyValues
EmailString = ""
datetscv = ""
datefsv = ""
timetsv = ""
timefsv = ""
mcNsv = ""


# Labels
customerID = Label(NUi, text ="CUSTOMER EMAIL_ID")
L2 = Label(NUi, text = "TIMESTAMP")
l12 = Label(NUi, text = "FROM")
l13 = Label(NUi, text = "to")


mcN = Label(NUi, text = "MCNumber")


# positions of labels
customerID.place(x = 10, y = 100)
L2.place(x = 10, y = 150)
l12.place(x = 105,y= 150)
l13.place(x = 105,y= 180 )
mcN.place(x = 10, y = 230)


# functions for labels
EnterID = Entry(NUi, bd =5, width = 50)
EnterID.insert(0,'username@xyz.de')

datestampfrom = Entry(NUi, bd =5, width =15)
datestampfrom.insert(0, 'YYYY-MM-DD')
timestamofrom = Entry(NUi, bd = 5, width = 15)
timestamofrom.insert(END, 'hh:mm:ss')
datestampto = Entry(NUi, bd =5, width = 15)
datestampto.insert(0, 'YYYY-MM-DD')
timestampto = Entry(NUi,bd = 5, width = 15 )
timestampto.insert(0,'hh:mm:ss')

mcNEnter = Entry(NUi, bd =5, width = 10)

EnterID.place(x = 150, y = 100)
datestampfrom.place(x = 150, y = 150)
datestampto.place(x= 150, y = 180)
timestamofrom.place(x = 265, y = 150 )
timestampto.place(x = 265, y = 180)
mcNEnter.place(x = 150, y = 230)


NUi.geometry("600x400+20+20")


def reportMessage():
    msg = messagebox.showinfo("","Click OK to proceed")
    global EnterIDsv, datestampfromsv, datestamptosv, mcNumbersv, timestampfromsv, timestamptosv
    EnterIDsv = EnterID.get()
    datestampfromsv = datestampfrom.get() + timestamofrom.get()
    datestamptosv = datestampto.get() + timestampto.get()
   #timestampfromsv = timestamofrom.get()
    #timestamptosv = timestampto.get()
    mcNumbersv = mcNEnter.get()
    NUi.destroy()


sReport = Button(NUi, text = "Request Report", command = reportMessage)
sReport.place(x = 200 , y = 300)
NUi.mainloop()

bad_chars = [':','-', ';', '']
def checkstrings(text):
 for i in bad_chars:
     text = text.replace(i,'')
 return text

EmailString = EnterIDsv
timefsv = checkstrings(datestampfromsv)
timetsv = checkstrings(datestamptosv)
mcNsv = mcNumbersv


print(EmailString, timefsv,timetsv, mcNsv)