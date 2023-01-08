from tkinter import *
from tkinter import ttk
from tkinter import Canvas
import pandas as pd
import numpy as np
#**************************************************************************************************************************#
pd.set_option('display.max_rows', 1000)
pd.set_option('display.max_columns', 1000)
#**************************************************************************************************************************#
def connect():
    global data
    global button_state
    button_state="DISABLED"
    _link=frame1_entry1.get()
    if(_link==""):
        frame4_textarea1.insert(END,"Nothing has been Entered")
        return False
    else:
        try:
            data=pd.read_excel(_link)
            frame4_textarea1.delete("1.0",END)
            frame4_textarea1.insert(END,"File Read Success")
            button_state="active"
            coldata1()
            coldata2()
            coldata3()
            coldata4()
            return True
        except:
            frame4_textarea1.delete("1.0",END)
            frame4_textarea1.insert(END,"File Read Error : File Link is not Valid")
            return False
#**************************************************************************************************************************#
def coldata1():
    global data
    frame2_dropdown1['values']=tuple(data.columns)
#**************************************************************************************************************************#
def coldata2():
    global data
    frame2_dropdown2['values']=tuple(data.columns)
#**************************************************************************************************************************#
def coldata3():
    global data
    frame2_dropdown3['values']=tuple(data.columns)
#**************************************************************************************************************************#
def coldata4():
    global data
    frame2_dropdown4['values']=tuple(data.columns)
#**************************************************************************************************************************#
def readdata():
    global data

    if(connect()==True):
        if(frame2_dropdown3.get()==""):
            frame4_textarea1.delete("1.0",END)
            frame4_textarea1.insert(END,"Choose Customer")
        else:
            w=str(frame2_dropdown1.get())
            x=str(frame2_dropdown2.get())
            y=str(frame2_dropdown3.get())
            z=str(frame2_dropdown4.get())


            df=data[[w,x,y,z]]
            frame5_textarea1.delete("1.0",END)
            frame5_textarea1.insert(END,df)
    else:
        frame4_textarea1.delete("1.0",END)
        frame4_textarea1.insert(END,"Cannot read from The File as File Connection is Invalid")

def fulldata():
    connect()
    frame5_textarea1.delete("1.0",END)
    frame5_textarea1.insert(END,data)
    frame4_textarea1.delete("1.0",END)
    frame4_textarea1.insert(END,"Excel Read Success.....")

#**************************************************************************************************************************#    
root=Tk()

#root.geometry('640x480')
#**************************************************************************************************************************#
# section Frame 1
frame1=LabelFrame(master=root,text="1. Communication Info")
frame1.grid(row=0,column=0,rowspan=2)

frame1_label1=Label(master=frame1,text="Excel File Address:")
frame1_label1.grid(row=0,column=1)

frame1_entry1=Entry(master=frame1,width=40)
frame1_entry1.grid(row=0,column=2)
frame1_entry1.insert(0,"Enter the Address of the Excel File")
#**************************************************************************************************************************#
# section Frame 2
#**************************************************************************************************************************#
frame2=LabelFrame(master=root,text="2: Product Info : Select Columns")
frame2.grid(row=2,column=0,rowspan=2)

frame2_label1=Label(master=frame2,text="Column 1 :")
frame2_label1.grid(row=0,column=0)

frame2_dropdown1=ttk.Combobox(master=frame2,width=50)
frame2_dropdown1.grid(row=0,column=1)

frame2_label2=Label(master=frame2,text="Column 2 :")
frame2_label2.grid(row=1,column=0)

frame2_dropdown2=ttk.Combobox(master=frame2,width=50)
frame2_dropdown2.grid(row=1,column=1)

frame2_label3=Label(master=frame2,text="Column 3 :")
frame2_label3.grid(row=2,column=0)

frame2_dropdown3=ttk.Combobox(master=frame2,width=50)
frame2_dropdown3.grid(row=2,column=1)


frame2_label4=Label(master=frame2,text="Column 4 :")
frame2_label4.grid(row=3,column=0)

frame2_dropdown4=ttk.Combobox(master=frame2,width=50)
frame2_dropdown4.grid(row=3,column=1)
#**************************************************************************************************************************#
#section Frame 3
#**************************************************************************************************************************#
frame3=LabelFrame(master=root,text="3: Connection/Programming :")
frame3.grid(row=4,column=0,rowspan=2)

frame3_button1=Button(master=frame3,text="Connect",command=connect)
frame3_button1.grid(row=0,column=0)

frame3_label1=Label(master=frame3,text="                ")
frame3_label1.grid(row=0,column=1)


frame3_button2=Button(master=frame3,text="Full Excel",command=fulldata)
frame3_button2.grid(row=0,column=2)

frame3_label2=Label(master=frame3,text="                ")
frame3_label2.grid(row=0,column=3)


frame3_button3=Button(master=frame3,text="Get Data : Selected Col",command=readdata)
frame3_button3.grid(row=0,column=4)
#**************************************************************************************************************************#
#section Frame 4
#**************************************************************************************************************************#
frame4=LabelFrame(master=root,text="4: Status")
frame4.grid(row=6,column=0)

frame4_textarea1=Text(master=frame4,height=1,width=50)
frame4_textarea1.grid(row=0,column=0)

#**************************************************************************************************************************#
#section Frame 5
#**************************************************************************************************************************#
frame5=LabelFrame(master=root,text="4: Data Acquired")
frame5.grid(row=7,column=0)

frame5_textarea1=Text(master=frame5,height=30,width=150)
frame5_textarea1.grid(row=0,column=0)
#**************************************************************************************************************************#

root.mainloop()