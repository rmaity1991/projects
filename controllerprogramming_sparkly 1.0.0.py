"""
    Initial Release : Rohit Maity
    Date : 21/08/2022
"""

from tkinter import *
import pandas as pd
import numpy as np
import os
#*************************************************************************************#
class sparkly:
    def __init__(self, Comm_Type, COM_Port, Baud_Rate, Bits, Parity,StopBits, Address):
        self.Comm_Type = Comm_Type
        self.COM_Port = COM_Port
        self.Baud_Rate = Baud_Rate
        self.Bits =Bits
        self.Parity = Parity
        self.StopBits = StopBits
        self.Address =Address
#*************************************************************************************#

def extractdata():
    Comm_Type = frame1_entry1.get()
    COM_Port = frame1_entry2.get()
    Baud_Rate = frame1_entry3.get()
    Bits =frame1_entry4.get()
    Parity = frame1_entry5.get()
    StopBits = frame1_entry6.get()
    Address =frame1_entry7.get()

    connection=sparkly(Comm_Type,COM_Port,Baud_Rate,Bits,Parity,StopBits,Address)

    command ="--connection " + connection.Comm_Type +','+ connection.COM_Port+','+connection.Baud_Rate + connection.Bits + connection.Parity +connection.StopBits +','+connection.Address
    
    # device read-data --connection Serial,COM4,1152008N2

    return command
#*************************************************************************************#
def deviceconnect():

    connection=extractdata();

    command = "sparkly device read-data " + connection
    os.system(command)
#*************************************************************************************#
def senddata():
    connection=extractdata()

    keys=frame3_entry1.get()
    values=frame3_entry2.get()

    keys_list=[x for x in keys.split(",")]
    value_list=[x for x in values.split(",")]

    message=""

    for index in range(len(keys_list)):
        data_message="{keys}={value}".format(keys=keys_list[index],value=value_list[index])
        message=message+" "+data_message
        
    print(message)

    command="sparkly parameters write --parameter-list "+ " " + message + " " +connection
    os.system(command)
#*************************************************************************************#
def download():
    connection=extractdata()

    path=frame2_entry1.get()

    command='sparkly app download --src ' + '"' + 'flash.bin' + '"' + ' ' + connection + ' '+'--working-directory' + ' '+ '"' + path + '"'

    os.system(command)
#*************************************************************************************#
def readdata():
    connection=extractdata()

    keys=frame4_entry1.get()

    keys_list=[x for x in keys.split(",")]

    message=""

    for index in range(len(keys_list)):
        data_message="{keys}".format(keys=keys_list[index])
        message=message+" "+data_message

    command="sparkly parameters read --parameter-list "+ " " + message + " " +connection
    os.system(command)



#****************************************************************************************#

root=Tk()
root.title("Mpx One Programming Tool")
# root.iconbitmap(r'logo.ico')
root.config(bg='#63B2FF')

#****************************************************************************************#

frame1=LabelFrame(master=root,text="1.1 Communication Info",bg='#9CC3D5',font='Tahoma',border=5)
frame1.grid(row=0,column=0)

frame1_label1=Label(master=frame1,text="COM Type :",bg='#9CC3D5',font='Tahoma')
frame1_label1.grid(row=0,column=0)

frame1_entry1=Entry(master=frame1,width=40,font='Tahoma')
frame1_entry1.grid(row=0,column=1)
frame1_entry1.insert(0,"Serial")


frame1_label2=Label(master=frame1,text="COM Port :",bg='#9CC3D5',font='Tahoma')
frame1_label2.grid(row=1,column=0)

frame1_entry2=Entry(master=frame1,width=40,font='Tahoma')
frame1_entry2.grid(row=1,column=1)
frame1_entry2.insert(0,"COM10")

frame1_label3=Label(master=frame1,text="Baud Rate :",bg='#9CC3D5',font='Tahoma')
frame1_label3.grid(row=2,column=0)

frame1_entry3=Entry(master=frame1,width=40,font='Tahoma')
frame1_entry3.grid(row=2,column=1)
frame1_entry3.insert(0,"19200")

frame1_label4=Label(master=frame1,text="Bits :",bg='#9CC3D5',font='Tahoma')
frame1_label4.grid(row=3,column=0)


frame1_entry4=Entry(master=frame1,width=40,font='Tahoma')
frame1_entry4.grid(row=3,column=1)
frame1_entry4.insert(0,"8")

frame1_label5=Label(master=frame1,text="Parity :",bg='#9CC3D5',font='Tahoma')
frame1_label5.grid(row=4,column=0)

frame1_entry5=Entry(master=frame1,width=40,font='Tahoma')
frame1_entry5.grid(row=4,column=1)
frame1_entry5.insert(0,"N")

frame1_label6=Label(master=frame1,text="Stop Bits :",bg='#9CC3D5',font='Tahoma')
frame1_label6.grid(row=5,column=0)

frame1_entry6=Entry(master=frame1,width=40,font='Tahoma')
frame1_entry6.grid(row=5,column=1)
frame1_entry6.insert(0,"2")

frame1_label7=Label(master=frame1,text="Device Address :",bg='#9CC3D5',font='Tahoma')
frame1_label7.grid(row=6,column=0)

frame1_entry7=Entry(master=frame1,width=40,font='Tahoma')
frame1_entry7.grid(row=6,column=1)
frame1_entry7.insert(0,"15")

frame1_button1=Button(master=frame1,text="Test Connection",command=deviceconnect,font='Tahoma')
frame1_button1.grid(row=7,column=2)

#************************************************************************************* #

frame2=LabelFrame(master=root,text="1.2 Download To Controller",bg='#9CC3D5',font='Tahoma',border=5)
frame2.grid(row=1,column=0)

frame2_label1=Label(master=frame2,text="Enter Path to Flash Bin :",bg='#9CC3D5',font='Tahoma')
frame2_label1.grid(row=0,column=0)


frame2_entry1=Entry(master=frame2,width=40,font='Tahoma')
frame2_entry1.grid(row=0,column=1)
frame2_entry1.insert(0,"Enter Path to Project flash.bin file")

frame2_button1=Button(master=frame2,text="Download",command=download,font='Tahoma')
frame2_button1.grid(row=1,column=2)


# *********************************************************************************** #

frame3=LabelFrame(master=root,text="1.3 Send Data To Controller",bg='#9CC3D5',font='Tahoma',border=5)
frame3.grid(row=2,column=0)

frame3_label1=Label(master=frame3,text="Controller Data Abbreviatios :",bg='#9CC3D5',font='Tahoma')
frame3_label1.grid(row=0,column=0)

frame3_entry1=Entry(master=frame3,width=40,font='Tahoma')
frame3_entry1.grid(row=0,column=1)
frame3_entry1.insert(0,"Enter Abbreviations Separated By Comma... Ex: H1,H2...")

frame3_label2=Label(master=frame3,text="Data Value for Abbreviatios :",bg='#9CC3D5',font='Tahoma')
frame3_label2.grid(row=1,column=0)

frame3_entry2=Entry(master=frame3,width=40,font='Tahoma')
frame3_entry2.grid(row=1,column=1)
frame3_entry2.insert(0,"Enter Value Separated By Comma... Ex: 12,23...")

frame3_button1=Button(master=frame3,text="Send Data",command=senddata,font='Tahoma')
frame3_button1.grid(row=2,column=2)

#*************************************************************************************#
frame4=LabelFrame(master=root,text="1.4 Read Data From Controller",bg='#9CC3D5',font='Tahoma',border=5)
frame4.grid(row=3,column=0)

frame4_label1=Label(master=frame4,text="Controller Data Abbreviatios :",bg='#9CC3D5',font='Tahoma')
frame4_label1.grid(row=0,column=0)

frame4_entry1=Entry(master=frame4,width=40,font='Tahoma')
frame4_entry1.grid(row=0,column=1)
frame4_entry1.insert(0,"Enter Abbreviations Separated By Comma... Ex: H1,H2...")

frame4_button1=Button(master=frame4,text="Read Data",command=readdata,font='Tahoma')
frame4_button1.grid(row=1,column=2)


#*************************************************************************************#

root.mainloop()