"""
    Initial Release : Rohit Maity
    Date : 27/11/2022
"""

from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter.filedialog import asksaveasfile
import openpyxl
import pandas as pd
import os
import serial.tools.list_ports
import numpy as np
#************************************************************************************* #
pd.set_option('display.max_rows', 1000)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)
#************************************************************************************* #

class sparkly:
    def __init__(self, Comm_Type, COM_Port, Baud_Rate, Bits, Parity,StopBits, Address):
        self.Comm_Type = Comm_Type
        self.COM_Port = COM_Port
        self.Baud_Rate = Baud_Rate
        self.Bits =Bits
        self.Parity = Parity
        self.StopBits = StopBits
        self.Address =Address

#************************************************************************************* #
def setrow(row_select):
    global enable_row
    global enable_col
    if (row_select):
        enable_row = ACTIVE
        enable_col = DISABLED
    else:
        enable_row = DISABLED
        enable_col = ACTIVE

#************************************************************************************* #
def connect():
    global data
    global read_excel
    global button_state
    button_state = "DISABLED"
    # _link = str(frame1_entry8.get())
    _link = filedialog.askopenfilename()
    print(_link)
    if (_link == ""):
        frame4_textarea1.insert(END, "Nothing has been Entered..................")
        read_excel=False
    else:
        try:
            data = pd.read_excel(_link)
            #print(data)
            frame4_textarea1.delete("1.0", END)
            frame4_textarea1.insert(END, "File Read Success................")
            button_state = "active"
            customer()
            category()
            casetype()
            parameter()
            read_excel=True
        except:
            frame4_textarea1.delete("1.0", END)
            frame4_textarea1.insert(END, "File Read Error : File Link is not Valid................")
            read_excel=False


#************************************************************************************* #
def customer():
    global data
    frame2_dropdown3['values'] = ('Dollar Tree', 'Publix', 'Aldi')

#************************************************************************************* #
def category():
    global data
    frame2_dropdown1['values'] = tuple(data.columns)

#************************************************************************************* #
def casetype():
    global data
    frame2_dropdown2['values'] = tuple(data.columns)


#************************************************************************************* #
def parameter():
    frame2_dropdown4['values'] = tuple(data['Name'])


#************************************************************************************* #
def readdatacolumn():
    global data
    global df
    global read_excel

    if (read_excel == True):
        if (frame2_dropdown3.get() == "" or frame2_dropdown2.get() == "" or frame2_dropdown1.get() == ""):
            frame4_textarea1.delete("1.0", END)
            frame4_textarea1.insert(END, "Choose All Columns for Data.................")
        else:
            x = str(frame2_dropdown3.get())
            y = str(frame2_dropdown1.get())
            z = str(frame2_dropdown2.get())

            frame4_textarea1.delete("1.0", END)
            frame4_textarea1.insert(END, "Data Read Success..............")

            df = data[['Name', 'Acronym', y, z, x]]
            frame5_textarea1.delete("1.0", END)
            frame5_textarea1.insert("1.0", df)
    else:
        frame4_textarea1.delete("1.0", END)
        frame4_textarea1.insert(END, "Cannot read from The File as File Connection is Invalid...............")


#*********************************************************************************************************** #
def readdatarow():
    global data
    global df
    global read_excel

    if (read_excel == True):
        if (frame2_dropdown4.get() == ""):
            frame4_textarea1.delete("1.0", END)
            frame4_textarea1.insert(END, "Choose a parameter for Data..............")
        else:
            x = str(frame2_dropdown4.get())

            frame4_textarea1.delete("1.0", END)
            frame4_textarea1.insert(END, "Data Read Success................")

            df = data[data['Name'] == x]
            frame5_textarea1.delete("1.0", END)
            frame5_textarea1.insert(END, df[
                ['Name', 'MinValue', 'MaxValue', 'InitValue', 'Category', 'Dollar Tree', 'Publix', 'Aldi']])
    else:
        frame4_textarea1.delete("1.0", END)
        frame4_textarea1.insert(END, "Cannot read from The File as File Connection is Invalid...............")


#*********************************************************************************************************** #
def getData():
    global data
    global df

    try:
        text_data=frame5_textarea1.get("1.0",END)
        fObj=open("customerdata.txt",mode='w')

        for var in text_data:
            fObj.write(var)
        fObj.close()

        wb=openpyxl.Workbook()
        sheet=wb.active
  
    
        fObj=open("customerdata.txt",mode='r')
        row=1
        for key,line in enumerate(fObj):
            line=line.split()
            if(key==0):
                line.insert(0,'Index')
            else:
                pass
                      
            range=len(line)
            for key,var in enumerate(line):
                sheet.cell(row=row,column=key+1).value=var
            row=row+1
        
        wb.save('customerdata.xlsx')
        path=os.path.abspath(str(os.path.dirname('customerdata.xlsx')))
        frame4_textarea1.delete("1.0", END)
        frame4_textarea1.insert(END, "Data Sent To Excel............")
        frame4_textarea1.insert(END, path)
    except:
        frame4_textarea1.delete("1.0", END)
        frame4_textarea1.insert(END, "Error in Sending Data to the Excel.....")


#********************************************************************************************** #

# [(keys1,value1),(keys2,value2),()]

def program():
    try:
        wb=pd.read_excel('customerdata.xlsx')
        keys=wb['Acronym']
        customer=str(frame2_dropdown3.get())
        values=wb[customer]
        keys.fillna('0',inplace=True)
        values.fillna('0',inplace=True)
        combined=pd.Series(zip(keys,values))
        message=""
        for item,value in combined:
            if item=='0' or value=='0':
                continue
            else:
                data_message="{keys}={value}".format(keys=item,value=value)
                message=message+" "+data_message
    
        frame4_textarea1.delete("1.0", END)
        frame4_textarea1.insert(END, "Data Extraction Success from customerdata............")
        senddata(message)
    except:
        frame4_textarea1.delete("1.0", END)
        frame4_textarea1.insert(END, "Data Extraction Failed from customerdata.............")


#*********************************************************************************************** #
def extractdata():
    Comm_Type = frame1_entry1.get()
    COM_Port = frame1_dropdown1.get()
    Baud_Rate = frame1_entry3.get()
    Bits =frame1_entry4.get()
    Parity = frame1_entry5.get()
    StopBits = frame1_entry6.get()
    Address =frame1_entry7.get()
    connection=sparkly(Comm_Type,COM_Port,Baud_Rate,Bits,Parity,StopBits,Address)
    command ="--connection " + connection.Comm_Type +','+ connection.COM_Port+','+connection.Baud_Rate + connection.Bits + connection.Parity +connection.StopBits +','+connection.Address
    return command


#*************************************************************************************#
def senddata(message):
    connection=extractdata()
    command="sparkly parameters write --parameter-list "+ " " + message + " " +connection
    os.system(command)
    frame4_textarea1.delete("1.0", END)
    frame4_textarea1.insert(END, "Data Sent To Controller")


#*************************************************************************************#
def deviceconnect():
    connection=extractdata();
    command = "sparkly device read-data " + connection
    os.system(command)

#*************************************************************************************#
def scanports():
    ports_info=serial.tools.list_ports.comports()
    ports=[]
    for item in ports_info:
        ports.append(item.name)   
    frame1_dropdown1['values']=tuple(ports)

   
#*************************************************************************************#
    
root = Tk()
# root.iconbitmap(r'logo.ico')
root.title('Mpxone Programming Tool')
root.configure(bg='#90AFC5')

#************************************************************************************* #
# Programming Window

# section Frame 1
frame1 = LabelFrame(master=root, text="1.1 Controller Communication Info",font=("ArialBlack",12),bg='#70AFC3',foreground='black',border=5)
frame1.grid(row=0, column=0, columnspan=5)

frame1_label1=Label(master=frame1,text="COM Type :",font=("Tahoma",10),bg='#70AFC3',foreground='black')
frame1_label1.grid(row=0,column=0)

frame1_entry1=Entry(master=frame1,width=40,font=("Tahoma",10),bg='#50AFC3')
frame1_entry1.grid(row=0,column=1)
frame1_entry1.insert(0,"Serial")


frame1_label2=Label(master=frame1,text="COM Port :",font=("Tahoma",10),bg='#70AFC3',foreground='black')
frame1_label2.grid(row=1,column=0)

# frame1_entry2=Entry(master=frame1,width=40,font=("Tahoma",10),bg='#50AFC3')
# frame1_entry2.grid(row=1,column=1)
# frame1_entry2.insert(0,"COM10")

frame1_dropdown1 = ttk.Combobox(master=frame1, width=50,font=("Tahoma",10))
frame1_dropdown1.grid(row=1, column=1)


frame1_button2 = Button(master=frame1, text="Scan For Ports", command=scanports,font=("Tahoma",10))
frame1_button2.grid(row=1, column=2)

frame1_label3=Label(master=frame1,text="Baud Rate :",font=("Tahoma",10),bg='#70AFC3',foreground='black')
frame1_label3.grid(row=2,column=0)

frame1_entry3=Entry(master=frame1,width=40,font=("Tahoma",10),bg='#50AFC3')
frame1_entry3.grid(row=2,column=1)
frame1_entry3.insert(0,"19200")

frame1_label4=Label(master=frame1,text="Bits :",font=("Tahoma",10),bg='#70AFC3',foreground='black')
frame1_label4.grid(row=3,column=0)


frame1_entry4=Entry(master=frame1,width=40,font=("Tahoma",10),bg='#50AFC3')
frame1_entry4.grid(row=3,column=1)
frame1_entry4.insert(0,"8")

frame1_label5=Label(master=frame1,text="Parity :",font=("Tahoma",10),bg='#70AFC3',foreground='black')
frame1_label5.grid(row=4,column=0)

frame1_entry5=Entry(master=frame1,width=40,font=("Tahoma",10),bg='#50AFC3')
frame1_entry5.grid(row=4,column=1)
frame1_entry5.insert(0,"N")

frame1_label6=Label(master=frame1,text="Stop Bits :",font=("Tahoma",10),bg='#70AFC3',foreground='black')
frame1_label6.grid(row=5,column=0)

frame1_entry6=Entry(master=frame1,width=40,font=("Tahoma",10),bg='#50AFC3')
frame1_entry6.grid(row=5,column=1)
frame1_entry6.insert(0,"2")

frame1_label7=Label(master=frame1,text="Device Address :",font=("Tahoma",10),bg='#70AFC3',foreground='black')
frame1_label7.grid(row=6,column=0)

frame1_entry7=Entry(master=frame1,width=40,font=("Tahoma",10),bg='#50AFC3')
frame1_entry7.grid(row=6,column=1)
frame1_entry7.insert(0,"15")

frame1_button1 = Button(master=frame1, text="Verify Controller Connection", command=deviceconnect,font=("Tahoma",10))
frame1_button1.grid(row=7, column=1)


frame1_label2 = Label(master=frame1, text="From Master Excel:",font=("Tahoma",10),bg='#70AFC3',foreground='black')
frame1_label2.grid(row=8, column=0)

# frame1_entry8 = Entry(master=frame1, width=40,font=("Tahoma",10),bg='#50AFC3')
# frame1_entry8.grid(row=8, column=1)
# frame1_entry8.insert(0, "Enter the Address of the Excel File")

frame1_button2 = Button(master=frame1, text="Browse Master Excel File", command=connect,font=("Tahoma",10))
frame1_button2.grid(row=8, column=2)

#************************************************************************************* #
# section Frame 2
frame2 = LabelFrame(master=root, text="2: Product Info : Name , Acronym Automatically Filled : Select Columns to Display",font=("ArialBlack",12),bg='#70AFC3',foreground='black',border=5)
frame2.grid(row=2, column=0, columnspan=5)


frame2_label1 = Label(master=frame2, text="Column 1 :",font=("Tahoma",10),bg='#70AFC3',foreground='black')
frame2_label1.grid(row=0, column=0)

frame2_dropdown1 = ttk.Combobox(master=frame2, width=50,font=("Tahoma",10))
frame2_dropdown1.grid(row=0, column=1)

frame2_label2 = Label(master=frame2, text="Column 2 :",font=("Tahoma",10),bg='#70AFC3',foreground='black')
frame2_label2.grid(row=1, column=0)

frame2_dropdown2 = ttk.Combobox(master=frame2, width=50,font=("Tahoma",10))
frame2_dropdown2.grid(row=1, column=1)

frame2_label3 = Label(master=frame2, text="Customer :",font=("Tahoma",10),bg='#70AFC3',foreground='black')
frame2_label3.grid(row=2, column=0)

frame2_dropdown3 = ttk.Combobox(master=frame2, width=50,font=("Tahoma",10))
frame2_dropdown3.grid(row=2, column=1)

frame2_label4 = Label(master=frame2, text="Parameter Acronym :",font=("Tahoma",10),bg='#70AFC3',foreground='black')
frame2_label4.grid(row=3, column=0)

frame2_dropdown4 = ttk.Combobox(master=frame2, width=50,font=("Tahoma",10))
frame2_dropdown4.grid(row=3, column=1)

frame2_button1 = Button(master=frame2, text="Get Data Column Wise", command=readdatacolumn,font=("Tahoma",10))
frame2_button1.grid(row=4, column=0)

frame2_label5 = Label(master=frame2, text="                ",bg='#70AFC3')
frame2_label5.grid(row=4, column=1)

frame2_button2 = Button(master=frame2, text="Get Data Row Wise", command=readdatarow,font=("Tahoma",10))
frame2_button2.grid(row=4, column=2)

frame2_label6 = Label(master=frame2, text="                ",bg='#70AFC3')
frame2_label6.grid(row=4, column=3)
#************************************************************************************* #
# section Frame 3

frame3 = LabelFrame(master=root, text="3: Send Data To Excel / Controller Programming :",font=("ArialBlack",12),bg='#70AFC3',foreground='black',border=5)
frame3.grid(row=4, column=0, columnspan=5)

frame3_button1 = Button(master=frame3, text="Get Current Data To Excel",command=getData,font=("Tahoma",10))
frame3_button1.grid(row=0, column=2)

frame3_label1 = Label(master=frame3, text="                ",bg='#70AFC3')
frame3_label1.grid(row=0, column=3)

frame3_button5 = Button(master=frame3, text="Program To Controller", command=program,font=("Tahoma",10))
frame3_button5.grid(row=0, column=8)
#************************************************************************************* #
# section Frame 4

frame4 = LabelFrame(master=root, text="4: Operation Status",font=("ArialBlack",12),bg='#70AFC3',foreground='black',border=5)
frame4.grid(row=6, column=0, columnspan=5)


frame4_textarea1 = Text(master=frame4, height=3, width=50,font=("Tahoma",10))
frame4_textarea1.grid(row=0, column=0)
#************************************************************************************* #
# section Frame 5

frame5 = LabelFrame(master=root, text="4: Data Acquired From Master Excel",font=("ArialBlack",12),bg='#90AFC5',foreground='black',border=5)
frame5.grid(row=7, column=0)

frame5_textarea1 = Text(master=frame5, height=30, width=150,bg='#90AFC5')
frame5_textarea1.grid(row=0, column=0)
#************************************************************************************* #
root.mainloop()