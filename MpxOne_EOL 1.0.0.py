"""
    Initial Release : Rohit Maity
    Date : 21/10/2022
"""

import time,os,openpyxl
from carelapi import Carel,Utility


master=True
controller_present=False
os.system('cls')
while master:
    os.system('cls')
    print("\t \t \t \t \tWelcome to Hill Pheonix - Dover Corporation - Carel EOL Testing \t \t \t \t ")
    print("\t \t \t \t \t***********************************************************\t \t \t \t \t")
    print("\t \t \t \t \t \t \t \t Main Menu \t \t \t \t ")
    print("\t \t \t \t \t***********************************************************\t \t \t \t \t")
    print("\t \t \t \t \t1 -   Define the Carel controller\t \t \t \t \t")
    print("\t \t \t \t \t2 -   Start Performing Tests\t \t \t \t \t")
    print("\t \t \t \t \t3 -   Get Last Test Results\t \t \t \t \t")
    print("\t \t \t \t \t4 -   Delete Defined Controller\t \t \t \t \t")
    print("\t \t \t \t \t5 -   Exit EOL Application\t \t \t \t \t")
    print("\t \t \t \t \tEnter your choice below:   \t \t \t \t \t")
    main_menu_option=int(input("\t \t \t \t \t"))
    time.sleep(1)
    
    if(main_menu_option==1):
        os.system('cls')
        print("\t \t \t \t \t***********************************************************\t \t \t \t \t")
        print("\t \t \t \t \t \t \t \t Controller Selection Menu \t \t \t \t ")
        print("\t \t \t \t \t***********************************************************\t \t \t \t \t")
        print("\t \t \t \t \t  Enter the name of the controller : \t \t \t \t ")
        print("\t \t \t \t \t  1 - 'mpxone' with 'Sparkly' \t \t \t \t ")
        print("\t \t \t \t \t  2 - 'mpxone' with 'Mod-Bus' \t \t \t \t ")
        print("\t \t \t \t \t  3 - 'cpco' with 'Serial' \t \t \t \t ")
        print("\t \t \t \t \t  4 - 'cpco' with 'Mod-Bus' \t \t \t \t ")
        print("\t \t \t \t \t  5 - 'pr300' with 'Serial' \t \t \t \t ")
        print("\t \t \t \t \t  6 - 'pr300' with 'Mod-Bus' \t \t \t \t ")
        print("\t \t \t \t \t  7 -  Exit to Main Menu \t \t \t \t ")
        option=int(input("\t \t \t \t \t"))
        if(option<1 and option>6):
            print("\t \t \t \t \tYour option is not listed, please try again\t \t \t \t \t")
        elif(option==7):
            pass
        else:
            controller=Carel(option)
            controller_present=True
            print(controller)

        time.sleep(1)


    if(main_menu_option==2):
        os.system('cls')
        submenu1=True
        if(controller_present==False):
            print("\t \t \t \t \tFirst Define the Controller....\t \t \t \t \t")
            print("\n \n \n")
            time.sleep(1)
        else:
            while submenu1:
                print("\t \t \t \t \t***********************************************************\t \t \t \t \t")
                print("\t \t \t \t \t \t \t \t Testing Menu \t \t \t \t ")
                print("\t \t \t \t \t***********************************************************\t \t \t \t \t")
                print("\t \t \t \t \t  Enter the test to perform of the controller : \t \t \t \t ")
                print("\t \t \t \t \t  1 - Test Connection \t \t \t \t ")
                print("\t \t \t \t \t  2 - Light Test \t \t \t \t ")
                print("\t \t \t \t \t  3 - EEV Test \t \t \t \t ")
                print("\t \t \t \t \t  4 - Perform All Tests \t \t \t \t ")
                print("\t \t \t \t \t  5 - Exit To Main Menu \t \t \t \t ")
                option=int(input("\t \t \t \t \t"))
                time.sleep(1)

                if(option==1):
                    controller._connect()
                elif(option==2):
                    controller._lightTest()
                elif(option==3):
                    controller._eevTest()
                elif(option==4):
                    for item in range(3):
                        controller._lightTest()
                elif(option==5):
                    submenu1=False


    if(main_menu_option==3):
        os.system('cls')
        wb=openpyxl.load_workbook('Test Results.xlsx')
        sheet=wb['Results']

        for item in range(1,sheet.max_row):
            print("\t \t \t \t \t {date} \t \t {controller} \t \t  {test} \t \t  {result} ".format(date=sheet.cell(row=item+1,column=1).value,controller=sheet.cell(row=item+1,column=2).value,test=sheet.cell(row=item+1,column=3).value,result=sheet.cell(row=item+1,column=4).value))

        time.sleep(5)

    if(main_menu_option==4):
        del controller

    if(main_menu_option==5):
        os.system('cls')
        print("\t \t \t \t \t  Bye... Come Back Soon..... \t \t \t \t ")
        print("\n \n \n")
        master=False
        time.sleep(1)
