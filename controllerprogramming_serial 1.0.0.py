"""
    Initial Release : Rohit Maity
    Date : 21/08/2022
"""
from tkinter import *
import pandas as pd
import numpy as np
import os
from pymodbus.client.sync import ModbusSerialClient as ModbusClient 

#**************************************************************************************************************************#
def deviceconnect():

    method = frame1_entry1.get()
    port = frame1_entry2.get()
    baudrate = int(frame1_entry3.get())
    bytes =int(frame1_entry4.get())
    parity = frame1_entry5.get()
    stopbits = int(frame1_entry6.get())
    deviceid =int(frame1_entry7.get())

    try:
        connection=ModbusClient(method = method, port=port, stopbits = stopbits, bytesize = bytes, parity = parity , baudrate= baudrate)
        connection.connect()
        if(connection.connect()):
            frame5_textarea1.delete("1.0", END)
            frame5_textarea1.insert(END, "Connection is Successfull..........")
            connection.close()
            return connection
        else:
            frame5_textarea1.delete("1.0", END)
            frame5_textarea1.insert(END, "Connection has Failed..........")
            connection.close()
    except:
            frame5_textarea1.delete("1.0", END)
            frame5_textarea1.insert(END, "Connection is Faulty and requires checking parameters..........")
            connection.close()

#**************************************************************************************************************************#
def coilsenddata():

    connection=deviceconnect()
    deviceid=int(frame1_entry7.get())

    keys=int(frame3_entry1.get())
    values=frame3_entry2.get()
    value=False

    if(values == "1"):
        value=True
    elif(values == "0"):
        value=False
    else:
        frame5_textarea1.delete("1.0", END)
        frame5_textarea1.insert(END, "Enter value 1 or 0 in values..........")

    connection.write_coil(keys,value,unit = deviceid) 
    connection.close()

#**************************************************************************************************************************#

def readcoildata():

    connection=deviceconnect()
    deviceid=int(frame1_entry7.get())

    keys=int(frame4_entry1.get())
    if(keys>300):
        frame5_textarea1.delete("1.0", END)
        frame5_textarea1.insert(END, "Enter Value less than 300..........")
    elif(keys<0):
        frame5_textarea1.delete("1.0", END)
        frame5_textarea1.insert(END, "Enter Value greater than 0..........")
    
    try:
        var=connection.read_coils(0x00,keys,unit=deviceid)
        inputcoils=var.bits
        var={}
        for i in range(keys):
            key=f"0x{i}"
            var[key]=inputcoils[i]

        result=""

        for i in var.keys():
            result=result+f"The value of coil address {i} is {var[i]} \n"

        frame6_textarea1.delete("1.0", END)
        frame6_textarea1.insert(END, result)
        connection.close()
    except:
        frame6_textarea1.delete("1.0", END)
        frame6_textarea1.insert(END, "Error in reading coils..........")

#**************************************************************************************************************************#
def readholdingregister():

    connection=deviceconnect()
    deviceid=int(frame1_entry7.get())

    try:
        var={}

        client=connection.read_holding_registers(0x00,100,unit=deviceid)
        holdingdata1=client.registers

        client=connection.read_holding_registers(0x101,100,unit=deviceid)
        holdingdata2=client.registers

        client=connection.read_holding_registers(0x201,100,unit=deviceid)
        holdingdata3=client.registers

        client=connection.read_holding_registers(0x301,100,unit=deviceid)
        holdingdata4=client.registers


       
        for i in range(100):
            key=f"0x{i}"
            var[key]=holdingdata1[i]
        for i in range(100):
            key=f"0x{i+100}"
            var[key]=holdingdata2[i]
        for i in range(100):
            key=f"0x{i+200}"
            var[key]=holdingdata3[i]
        for i in range(100):
            key=f"0x{i+300}"
            var[key]=holdingdata4[i]


        result=""

        for i in var.keys():
            result=result+f"The value of holding address {i} is {var[i]} \n"

        frame6_textarea1.delete("1.0", END)
        frame6_textarea1.insert(END, result)
        connection.close()
    except:
        frame6_textarea1.delete("1.0", END)
        frame6_textarea1.insert(END, "Error in reading holding registers..........")

#**************************************************************************************************************************#
def readinputregister():

    connection=deviceconnect()
    deviceid=int(frame1_entry7.get())

    try:
        var={}

        client=connection.read_input_registers(0x00,100,unit=deviceid)
        inputdata1=client.registers

        client=connection.read_input_registers(0x101,100,unit=deviceid)
        inputdata2=client.registers

        client=connection.read_input_registers(0x201,100,unit=deviceid)
        inputdata3=client.registers

        client=connection.read_input_registers(0x301,100,unit=deviceid)
        inputdata4=client.registers


       
        for i in range(100):
            key=f"0x{i}"
            var[key]=inputdata1[i]
        for i in range(100):
            key=f"0x{i+100}"
            var[key]=inputdata2[i]
        for i in range(100):
            key=f"0x{i+200}"
            var[key]=inputdata3[i]
        for i in range(100):
            key=f"0x{i+300}"
            var[key]=inputdata4[i]


        result=""

        for i in var.keys():
            result=result+f"The value of input address {i} is {var[i]} \n"

        frame6_textarea1.delete("1.0", END)
        frame6_textarea1.insert(END, result)
        connection.close()
    except:
        frame6_textarea1.delete("1.0", END)
        frame6_textarea1.insert(END, "Error in reading input registers..........")
#**************************************************************************************************************************#

def holdingsenddata():

    connection=deviceconnect()
    deviceid=int(frame1_entry7.get())

    keys=int(frame5_3_entry1.get())
    values=int(frame5_3_entry2.get())

    connection.write_register(keys,values,unit = deviceid) 
    connection.close()

#**************************************************************************************************************************#
root=Tk()
root.title("Serial RTU Modbus Programming Tool")
root.iconbitmap(r'favicon.ico')
root.geometry('1200x950')
#**************************************************************************************************************************#

frame1=LabelFrame(master=root,text="1.1 Communication Info")
frame1.grid(row=0,column=0)

frame1_label1=Label(master=frame1,text="COM Type :")
frame1_label1.grid(row=0,column=0)

frame1_entry1=Entry(master=frame1,width=40)
frame1_entry1.grid(row=0,column=1)
frame1_entry1.insert(0,"rtu")


frame1_label2=Label(master=frame1,text="COM Port :")
frame1_label2.grid(row=1,column=0)

frame1_entry2=Entry(master=frame1,width=40)
frame1_entry2.grid(row=1,column=1)
frame1_entry2.insert(0,"COM11")

frame1_label3=Label(master=frame1,text="Baud Rate :")
frame1_label3.grid(row=2,column=0)

frame1_entry3=Entry(master=frame1,width=40)
frame1_entry3.grid(row=2,column=1)
frame1_entry3.insert(0,"19200")

frame1_label4=Label(master=frame1,text="Byte Size :")
frame1_label4.grid(row=3,column=0)


frame1_entry4=Entry(master=frame1,width=40)
frame1_entry4.grid(row=3,column=1)
frame1_entry4.insert(0,"8")

frame1_label5=Label(master=frame1,text="Parity :")
frame1_label5.grid(row=4,column=0)

frame1_entry5=Entry(master=frame1,width=40)
frame1_entry5.grid(row=4,column=1)
frame1_entry5.insert(0,"N")

frame1_label6=Label(master=frame1,text="Stop Bits :")
frame1_label6.grid(row=5,column=0)

frame1_entry6=Entry(master=frame1,width=40)
frame1_entry6.grid(row=5,column=1)
frame1_entry6.insert(0,"2")

frame1_label7=Label(master=frame1,text="Device Address :")
frame1_label7.grid(row=6,column=0)

frame1_entry7=Entry(master=frame1,width=40)
frame1_entry7.grid(row=6,column=1)
frame1_entry7.insert(0,"199")

frame1_button1=Button(master=frame1,text="Test Connection",command=deviceconnect)
frame1_button1.grid(row=7,column=0)


#**************************************************************************************************************************#

frame3=LabelFrame(master=root,text="1.2 Send Coils Data to Controller")
frame3.grid(row=1,column=0)

frame3_label1=Label(master=frame3,text="Input Coils Address :")
frame3_label1.grid(row=0,column=0)

frame3_entry1=Entry(master=frame3,width=40)
frame3_entry1.grid(row=0,column=1)
frame3_entry1.insert(0,"Enter single coil address")

frame3_label2=Label(master=frame3,text="Data Value:")
frame3_label2.grid(row=1,column=0)

frame3_entry2=Entry(master=frame3,width=40)
frame3_entry2.grid(row=1,column=1)
frame3_entry2.insert(0,"Ex: 1")

frame3_button1=Button(master=frame3,text="Write Coil Data",command=coilsenddata)
frame3_button1.grid(row=2,column=0)

#**************************************************************************************************************************#
frame4=LabelFrame(master=root,text="1.3 Read Coil Data From Controller")
frame4.grid(row=2,column=0)

frame4_label1=Label(master=frame4,text="Enter no of coils to read")
frame4_label1.grid(row=0,column=0)

frame4_entry1=Entry(master=frame4,width=40)
frame4_entry1.grid(row=0,column=1)
frame4_entry1.insert(0,"Ex: 0,10,20,50...300")

frame4_button1=Button(master=frame4,text="Read Coil Data",command=readcoildata)
frame4_button1.grid(row=1,column=0)


#**************************************************************************************************************************#

frame5_1=LabelFrame(master=root,text="1.4 Read Input Data From Controller")
frame5_1.grid(row=3,column=0)

frame5_1_label1=Label(master=frame5_1,text="Click to Read the Input Registers")
frame5_1_label1.grid(row=0,column=0)

frame5_1_button1=Button(master=frame5_1,text="Read Input Data",command=readinputregister)
frame5_1_button1.grid(row=0,column=1)


#**************************************************************************************************************************#

frame5_2=LabelFrame(master=root,text="1.5 Read Holding Data From Controller")
frame5_2.grid(row=4,column=0)

frame5_2_label1=Label(master=frame5_2,text="Click to Read the Input Registers")
frame5_2_label1.grid(row=0,column=0)

frame5_2_button1=Button(master=frame5_2,text="Read Input Data",command=readholdingregister)
frame5_2_button1.grid(row=0,column=1)
#**************************************************************************************************************************#
frame5_3=LabelFrame(master=root,text="1.6 Send Register Data to Controller")
frame5_3.grid(row=5,column=0)

frame5_3_label1=Label(master=frame5_3,text="Holding Register Address :")
frame5_3_label1.grid(row=0,column=0)

frame5_3_entry1=Entry(master=frame5_3,width=40)
frame5_3_entry1.grid(row=0,column=1)
frame5_3_entry1.insert(0,"Enter single register address")

frame5_3_label2=Label(master=frame5_3,text="Data Value:")
frame5_3_label2.grid(row=1,column=0)

frame5_3_entry2=Entry(master=frame5_3,width=40)
frame5_3_entry2.grid(row=1,column=1)
frame5_3_entry2.insert(0,"Ex: 36")

frame5_3_button1=Button(master=frame5_3,text="Write Data",command=holdingsenddata)
frame5_3_button1.grid(row=2,column=0)

#**************************************************************************************************************************#
frame5 = LabelFrame(master=root, text="1.7: Status")
frame5.grid(row=0,column=1)


frame5_textarea1 = Text(master=frame5, height=3, width=50)
frame5_textarea1.grid(row=0, column=0)

#**************************************************************************************************************************#

frame6 = LabelFrame(master=root, text="1.8: Data From Controller")
frame6.grid(row=1, column=1,columnspan=5)


frame6_textarea1 = Text(master=frame6, height=30, width=100)
frame6_textarea1.grid(row=0, column=0)
#**************************************************************************************************************************#
root.mainloop()