import pandas as pd
import math
# import xlwings as xw
#import np
import numpy 
from tkinter import *
from tkinter.ttk import *
from termcolor import colored
import sys
import pickle
from openpyxl import load_workbook

#-------------GUI WINDOW OPEN------------------------------------------------------------------------------
# root = Tk()

#-----FUNCTIONS-----------------------
#Refrigerant properties function:
#   ref: refrigerant
#   temp: Temperature
#   phase: Phase of fluid (Refrigerant) either Vapor/Liquid
def SatRefProp(ref,temp,phase):

    #Selecting the data table associated w/Refrigerant
    temp = round(temp)

    file_name = ref + '.Sat.Refprop'
    inFile1 = file_name
    fd1 = open(inFile1, 'rb')
    data = pickle.load(fd1)

    #Pulling Refrigerant Properties associated w/Temperature input
    tempR=""
    n=0
    numRows=len(data.loc[:, 'Temperature (F)'])
    for n in range(numRows):
        tempR=data.loc[n]['Temperature (F)']
        if tempR == temp:
            if phase == "Vapor":
                Press = data.loc[n]['Pressure (v,psia)']
                Density = data.loc[n]['Density (v, lbm/ft3)']
                Enthalpy = data.loc[n]['Enthalpy (v, Btu/lbm)']
                Dyn_Viscosity = data.loc[n]['Viscosity (v, lbm/ft*s)']
                Entropy = data.loc[n]['Entropy (v,BTU/lb*R)']
            if phase == "Liquid":
                Press = data.loc[n]['Pressure (l,psia)']
                Density = data.loc[n]['Density (l, lbm/ft3)']
                Enthalpy = data.loc[n]['Enthalpy (l, Btu/lbm)']
                Dyn_Viscosity = data.loc[n]['Viscosity (l, lbm/ft*s)']
                Entropy = data.loc[n]['Entropy (l,BTU/lb*R)']
            break
    if tempR!=temp:
        print(colored('Error - Temperature not in Data Tables', 'magenta'))
        sys.exit(1)
    return [Density,Enthalpy,Dyn_Viscosity,Press,Entropy]

#   Translated R&D tool for approximate refrigerant properties without storing refrigerant data
#   temp: temperature of CO2
#   press: pressure of CO2
def CO2_RefProp(temp, press):

    # Converting gauge pressure to atmospheric
    apress = press + 14.696

    # Using best fit formula for estimating saturated pressure from entered temperature
    if temp > 87:
        Sat_press = -9.97523E-10 * (temp ** 6) + 4.36367E-7 * (temp ** 5) - 4.52612E-5 * (temp ** 4) - 0.006323973 * \
                    (temp ** 3) + 1.687573758 * (temp ** 2) - 112.9195194 * temp + 3111.645388
    if temp <= 87:
        Sat_press = 3.431038E-11*(temp**6)-2.47282E-10*(temp**5)-7.41582E-8*(temp**4)+8.51149E-5*(temp**3)\
                    +0.031888303*(temp**2)+5.139252802*temp+290.9911272
    # Using best fit formula for estimating saturated temperature from entered pressure
    if apress >= 1070:
        Sat_temp = 5.98214520736336E-8*(apress**3)-0.000187842747608279*(apress**2)+0.273042293826637*apress-62.6596042187963
    if apress < 1070:
        # -2.83555168910717E-16*P1^6+1.2475949082525E-12*P1^5-2.2800927749137E-09*P1^4+2.25789866952818E-06*P1^3
        # -0.00135632711586607*P1^2+0.601795717807189*P1-104.874229892305
        Sat_temp = -2.83555168910717E-16*(apress**6)+1.2475949082525E-12*(apress**5)-2.2800927749137E-9*(apress**4)\
                   +2.25789866952818E-6*(apress**3)-0.00135632711586607*(apress**2)+0.601795717807189*apress\
                   -104.874229892305
    temp_diff = Sat_temp - temp
    TD = abs(temp_diff)
    presstemp = [press, temp]
    #triple point state properties
    tp_index = [[1060, 88], [1059, 88], [1058, 88], [1057, 88]]
    tp_enthal = [135.787, 137.8915, 149.958, 152.8604]
    tp_dens = [34.135, 32.626, 25.366, 23.922]
    tp_entrop = [0.3295, 0.33339, 0.35544, 0.36075]
    tp_visc = [2.65157E-5, 2.51149E-5, 1.9374E-5, 1.84118E-5]
    tp_matrix = numpy.array([tp_index,
                 tp_enthal,
                 tp_dens,
                 tp_entrop,
                 tp_visc])
    n = 0
    for tp in tp_index:
        if tp == presstemp:
            Density = tp_matrix[2][n]
            Enthalpy = tp_matrix[1][n]
            Dyn_Viscosity = tp_matrix[4][n]
            Entropy = tp_matrix[3][n]
            return [Density, Enthalpy, Dyn_Viscosity, press, Entropy, temp]
        n = n + 1

    def cubicformula(const, apress):
        p6 = apress ** 6
        p5 = apress ** 5
        p4 = apress ** 4
        p3 = apress ** 3
        p2 = apress ** 2
        val = p6 * const.loc['c6'] + p5 * const.loc['c5'] + p4 * const.loc['c4'] + p3 * const.loc['c3'] \
              + p2 * const.loc['c2'] + apress * const.loc['c1'] + const.loc['c0']
        return val

    # Data retrieval
    inFile1 = 'ent.const'
    fd1 = open(inFile1, 'rb')
    ent_data = pickle.load(fd1)

    inFile2 = 'dens.const'
    fd2 = open(inFile2, 'rb')
    dens_data = pickle.load(fd2)

    inFile3 = 'entrop.const'
    fd3 = open(inFile3, 'rb')
    entrop_data = pickle.load(fd3)

    inFile4 = 'visc.const'
    fd4 = open(inFile4, 'rb')
    visc_data = pickle.load(fd4)

    def Form_Interp(state, data):
        if state == 'Sub-Sub':
            const_arr = data.loc[0:9, ['TD', 'c0', 'c1', 'c2', 'c3', 'c4', 'c5', 'c6']]
            const = const_arr.loc[0, :]
            AA4 = cubicformula(const, apress)
            const = const_arr.loc[1, :]
            AA6 = cubicformula(const, apress)
            const = const_arr.loc[2, :]
            AA8 = cubicformula(const, apress)
            const = const_arr.loc[3, :]
            AA10 = cubicformula(const, apress)
            const = const_arr.loc[4, :]
            AA12 = cubicformula(const, apress)
            const = const_arr.loc[5, :]
            AA14 = cubicformula(const, apress)
            const = const_arr.loc[6, :]
            AA16 = cubicformula(const, apress)
            const = const_arr.loc[7, :]
            AA18 = cubicformula(const, apress)
            const = const_arr.loc[8, :]
            AA20 = cubicformula(const, apress)
            const = const_arr.loc[9, :]
            AA22 = cubicformula(const, apress)

            if TD == 0:
                x = AA4
            if 0 < TD < 0.2:
                x = AA6 - (AA6 - AA8) / (0.2 - 1) * (0.2 - TD)
            if TD == 0.2:
                x = AA6
            if 0.2 < TD < 1:
                x = AA6 - (AA6 - AA8) / (0.2 - 1) * (0.2 - TD)
            if TD == 1:
                x = AA8
            if 1 < TD < 3:
                x = AA8 - (AA8 - AA10) / (1 - 3) * (1 - TD)
            if TD == 3:
                x = AA10
            if 3 < TD < 5:
                x = AA10 - (AA10 - AA12) / (3 - 5) * (3 - TD)
            if TD == 5:
                x = AA12
            if 5 < TD < 10:
                x = AA12 - (AA12 - AA14) / (5 - 10) * (5 - TD)
            if TD == 10:
                x = AA14
            if 10 < TD < 25:
                x = AA14 - (AA14 - AA16) / (10 - 25) * (10 - TD)
            if TD == 25:
                x = AA16
            if 25 < TD < 42:
                x = AA16 - (AA16 - AA18) / (25 - 42) * (25 - TD)
            if TD == 42:
                x = AA18
            if 42 < TD < 70:
                x = AA18 - (AA18 - AA20) / (42 - 70) * (42 - TD)
            if TD == 70:
                x = AA20
            if 70 < TD < 100:
                x = AA20 - (AA20 - AA22) / (70 - 100) * (70 - TD)
            if TD == 100:
                x = AA22
            if TD > 100:
                print(colored('Error - Temperature not in Data Tables: Fluid subcooled lower than 100°F', 'magenta'))
                sys.exit(1)

        if state == 'Sub-Super':
            const_arr = data.loc[11:21, ['TD', 'c0', 'c1', 'c2', 'c3', 'c4', 'c5', 'c6']]
            const = const_arr.loc[11, :]
            AA4 = cubicformula(const, apress)
            const = const_arr.loc[12, :]
            AA6 = cubicformula(const, apress)
            const = const_arr.loc[13, :]
            AA8 = cubicformula(const, apress)
            const = const_arr.loc[14, :]
            AA10 = cubicformula(const, apress)
            const = const_arr.loc[15, :]
            AA12 = cubicformula(const, apress)
            const = const_arr.loc[16, :]
            AA14 = cubicformula(const, apress)
            const = const_arr.loc[17, :]
            AA16 = cubicformula(const, apress)
            const = const_arr.loc[18, :]
            AA18 = cubicformula(const, apress)
            const = const_arr.loc[19, :]
            AA20 = cubicformula(const, apress)
            const = const_arr.loc[20, :]
            AA22 = cubicformula(const, apress)
            const = const_arr.loc[21, :]
            AA24 = cubicformula(const, apress)
            if TD == 0:
                x = AA4
            if 0 < TD < 0.2:
                x = AA6 - (AA6 - AA8) / (0.2 - 1) * (0.2 - TD)
            if TD == 0.2:
                x = AA6
            if 0.2 < TD < 1:
                x = AA6 - (AA6 - AA8) / (0.2 - 1) * (0.2 - TD)
            if TD == 1:
                x = AA8
            if 1 < TD < 3:
                x = AA8 - (AA8 - AA10) / (1 - 3) * (1 - TD)
            if TD == 3:
                x = AA10
            if 3 < TD < 5:
                x = AA10 - (AA10 - AA12) / (3 - 5) * (3 - TD)
            if TD == 5:
                x = AA12
            if 5 < TD < 10:
                x = AA12 - (AA12 - AA14) / (5 - 10) * (5 - TD)
            if TD == 10:
                x = AA14
            if 10 < TD < 25:
                x = AA14 - (AA14 - AA16) / (10 - 25) * (10 - TD)
            if TD == 25:
                x = AA16
            if 25 < TD < 42:
                x = AA16 - (AA16 - AA18) / (25 - 42) * (25 - TD)
            if TD == 42:
                x = AA18
            if 42 < TD < 70:
                x = AA18 - (AA18 - AA20) / (42 - 70) * (42 - TD)
            if TD == 70:
                x = AA20
            if 70 < TD < 100:
                x = AA20 - (AA20 - AA22) / (70 - 100) * (70 - TD)
            if TD == 100:
                x = AA22
            if 100 < TD < 200:
                x = AA22 - (AA22 - AA24) / (100 - 200) * (100 - TD)
            if TD == 200:
                x = AA24
            if TD > 200:
                print(colored('Error - Temperature not in Data Tables: Fluid superheated higher than 200°F', 'magenta'))
                sys.exit(1)
        if state == 'Super-Sub':
            const_arr = data.loc[23:33, ['TD', 'c0', 'c1', 'c2', 'c3', 'c4', 'c5', 'c6']]
            const = const_arr.loc[23, :]
            AA4 = cubicformula(const, apress)
            const = const_arr.loc[24, :]
            AA6 = cubicformula(const, apress)
            const = const_arr.loc[25, :]
            AA8 = cubicformula(const, apress)
            const = const_arr.loc[26, :]
            AA10 = cubicformula(const, apress)
            const = const_arr.loc[27, :]
            AA12 = cubicformula(const, apress)
            const = const_arr.loc[28, :]
            AA14 = cubicformula(const, apress)
            const = const_arr.loc[29, :]
            AA16 = cubicformula(const, apress)
            const = const_arr.loc[30, :]
            AA18 = cubicformula(const, apress)
            const = const_arr.loc[31, :]
            AA20 = cubicformula(const, apress)
            const = const_arr.loc[32, :]
            AA22 = cubicformula(const, apress)
            const = const_arr.loc[33, :]
            AA24 = cubicformula(const, apress)

            if TD == 0:
                x = AA4
            if 0 < TD < 0.2:
                x = AA4 - (AA4 - AA6) / (0.2 - 1) * (0.2 - TD)
            if TD == 0.2:
                x = AA6
            if 0.2 < TD < 1:
                x = AA6 - (AA6 - AA8) / (0.2 - 1) * (0.2 - TD)
            if TD == 1:
                x = AA8
            if 1 < TD < 3:
                x = AA8 - (AA8 - AA10) / (1 - 3) * (1 - TD)
            if TD == 3:
                x = AA10
            if 3 < TD < 5:
                x = AA10 - (AA10 - AA12) / (3 - 5) * (3 - TD)
            if TD == 5:
                x = AA12
            if 5 < TD < 10:
                x = AA12 - (AA12 - AA14) / (5 - 10) * (5 - TD)
            if TD == 10:
                x = AA14
            if 10 < TD < 17:
                x = AA14 - (AA14 - AA16) / (10 - 17) * (10 - TD)
            if TD == 17:
                x = AA16
            if 17 < TD < 25:
                x = AA16 - (AA16 - AA18) / (17 - 25) * (17 - TD)
            if TD == 25:
                x = AA18
            if 25 < TD < 42:
                x = AA18 - (AA18 - AA20) / (25 - 42) * (25 - TD)
            if TD == 42:
                x = AA20
            if 42 < TD < 70:
                x = AA20 - (AA20 - AA22) / (42 - 70) * (42 - TD)
            if TD == 70:
                x = AA22
            if 70 < TD < 100:
                x = AA22 - (AA22 - AA24) / (70 - 100) * (70 - TD)
            if TD == 100:
                x = AA24
            if TD > 100:
                print(colored('Error - Temperature not in Data Tables: Fluid subcooled lower than 100°F', 'magenta'))
                sys.exit(1)
        if state == 'Super-Super':
            const_arr = data.loc[35:46, ['TD', 'c0', 'c1', 'c2', 'c3', 'c4', 'c5', 'c6']]
            const = const_arr.loc[35, :]
            AA4 = cubicformula(const, apress)
            const = const_arr.loc[36, :]
            AA6 = cubicformula(const, apress)
            const = const_arr.loc[37, :]
            AA8 = cubicformula(const, apress)
            const = const_arr.loc[38, :]
            AA10 = cubicformula(const, apress)
            const = const_arr.loc[39, :]
            AA12 = cubicformula(const, apress)
            const = const_arr.loc[40, :]
            AA14 = cubicformula(const, apress)
            const = const_arr.loc[41, :]
            AA16 = cubicformula(const, apress)
            const = const_arr.loc[42, :]
            AA18 = cubicformula(const, apress)
            const = const_arr.loc[43, :]
            AA20 = cubicformula(const, apress)
            const = const_arr.loc[44, :]
            AA22 = cubicformula(const, apress)
            const = const_arr.loc[45, :]
            AA24 = cubicformula(const, apress)
            const = const_arr.loc[46, :]
            AA26 = cubicformula(const, apress)

            if TD == 0:
                x = AA4
            if 0 < TD < 0.2:
                x = AA4 - (AA4 - AA6) / (0.2 - 1) * (0.2 - TD)
            if TD == 0.2:
                x = AA6
            if 0.2 < TD < 1:
                x = AA6 - (AA6 - AA8) / (0.2 - 1) * (0.2 - TD)
            if TD == 1:
                x = AA8
            if 1 < TD < 3:
                x = AA8 - (AA8 - AA10) / (1 - 3) * (1 - TD)
            if TD == 3:
                x = AA10
            if 3 < TD < 5:
                x = AA10 - (AA10 - AA12) / (3 - 5) * (3 - TD)
            if TD == 5:
                x = AA12
            if 5 < TD < 10:
                x = AA12 - (AA12 - AA14) / (5 - 10) * (5 - TD)
            if TD == 10:
                x = AA14
            if 10 < TD < 17:
                x = AA14 - (AA14 - AA16) / (10 - 17) * (10 - TD)
            if TD == 17:
                x = AA16
            if 17 < TD < 25:
                x = AA16 - (AA16 - AA18) / (17 - 25) * (17 - TD)
            if TD == 25:
                x = AA18
            if 25 < TD < 42:
                x = AA18 - (AA18 - AA20) / (25 - 42) * (25 - TD)
            if TD == 42:
                x = AA20
            if 42 < TD < 70:
                x = AA20 - (AA20 - AA22) / (42 - 70) * (42 - TD)
            if TD == 70:
                x = AA22
            if 70 < TD < 100:
                x = AA22 - (AA22 - AA24) / (70 - 100) * (70 - TD)
            if TD == 100:
                x = AA24
            if 100 < TD < 200:
                x = AA24 - (AA24 - AA26) / (100 - 200) * (100 - TD)
            if TD == 200:
                x = AA26
            if TD > 200:
                print(colored('Error - Temperature not in Data Tables: Fluid superheated more than 200°F', 'magenta'))
                sys.exit(1)
        prop = x
        return prop

    ############################ Excel function rewriting ######################################
    # ----------------------------------Enthalpy-------------------------------------------------
    LL_enthal = -0.00504938891240354*(temp**4)+1.21490260161298*(temp**3)-82.1299934187415*(temp**2)+111981.691033475
    HL_enthal = 0.00155176441200192*(temp**4)-0.391712547846896*(temp**3)+27.8737873138122*(temp**2)-40929.4206991088
    UL_enthal = 0.00398182462245312*(temp**4)-1.46277620182009*(temp**3)+201.333562847417*(temp**2)\
                -12293.47451037443*temp+281861.558850296

    ## ---------------------------Enthalpy calc------------------------------------------------
    if (temp_diff > 0 and apress < 1080) or (87 < temp < 93 and LL_enthal < apress < UL_enthal):
        state = 'Sub-Sub'
    if temp_diff < 0 and apress < 1080:
        state = 'Sub-Super'
    if temp_diff > 0 and apress > 1079:
        state = 'Super-Sub'
    if temp_diff < 0 and apress > 1079:
        state = 'Super-Super'
    Enthalpy = Form_Interp(state, ent_data)
    # ------------------------------------------------------------------------------------

    ## ----------------------------------Density------------------------------------------------------
    LL_dens = -0.0927682155040239*(temp**4)+33.2619568929506*(temp**3)-4470.52056334959*(temp**2)+266955.451655175*temp\
              -5975114.63284514
    HL_dens = -0.000172667399048802*(temp**4)+0.022552854582826*(temp**3)-4085.53103449557
    UL_dens = -0.00625237990181588*(temp**4)+2.28780863805437*(temp**3)-313.852077401611*(temp**2)\
                +19143.4321396295*temp-437165.388831531
    ## -----------------------------Density Calc-----------------------------------------------------
    if (temp_diff > 0 and apress < 1080) or (87 < temp < 94 and LL_dens <= apress < UL_dens) or \
            (93 < temp < 97 and HL_dens < apress < UL_dens):
        state = 'Sub-Sub'
    if temp_diff < 0 and apress < 1080:
        state = 'Sub-Super'
    if temp_diff > 0 and apress > 1079:
        state = 'Super-Sub'
    if temp_diff < 0 and apress > 1079:
        state = 'Super-Super'
    Density = Form_Interp(state, dens_data)
    # -------------------------------------------------------------------------------------------------

    ## ----------------------------------Entropy-----------------------------------------------------
    LL_entrop = -0.00501744203146906 * (temp ** 4) + 1.20543088173881 * (temp ** 3) - 81.3688890347245 * (temp ** 2) + 110626.422397475
    HL_entrop = 0.000942047411156232 * (temp ** 4) - 0.237469188601317 * (temp ** 3) + 16.9007836812499 * (temp ** 2) - 24480.5814798924
    UL_entrop = 0.00207578166046328 * (temp ** 4) - 0.760081696500258 * (temp ** 3) + 104.222579977113 * (temp ** 2) - 6331.02649615714 * temp + 144631.421299056
    ## --------------------------------Entropy Calc----------------------------------------------------
    if (temp_diff > 0 and apress < 1080) or (87 < temp_diff < 93 and LL_entrop < apress < UL_entrop) or \
            (92 < apress < 99 and HL_entrop < apress < UL_entrop):
        state = 'Sub-Sub'
    if temp_diff < 0 and apress < 1080:
        state = 'Sub-Super'
    if temp_diff > 0 and apress > 1079:
        state = 'Super-Sub'
    if temp_diff < 0 and apress > 1079:
        state = 'Super-Super'
    Entropy = Form_Interp(state, entrop_data)
    # ---------------------------------------------------------------------------------------------------

    # -----------------------------------Dynamic Viscosity----------------------------------------------
    LL_visc = -0.00596786557796885 * (temp ** 4) + 1.43569943328058 * (temp ** 3) - 97.0558007275146 * (temp ** 2) \
              + 132180.865017821
    HL_visc = 0.00155176441200192 * (temp ** 4) - 0.391712547846896 * (temp ** 3) + 27.8737873138122 * (temp ** 2) \
              - 40929.4206991088
    UL_visc = 0.014785265841317 * (temp ** 4) - 5.51650593443126 * (temp ** 3) + 771.42723918422 * (temp ** 2) \
              - 47907.1781621823 * temp + 1115687.90008923
    ## ---------------------------------Dynamic Viscosity Calc--------------------------------------
    # (AND(B5-D4>0,P1<1056),AND(D4>87,D4<93,DB55<P1,P1<DD55),AND(D4>92,D4<99,DC55<P1,P1<DD55))
    if (temp_diff > 0 and apress < 1056) or (87 < temp < 93 and LL_visc < apress < UL_visc) or \
        (92 < temp < 99 and HL_visc < apress < UL_visc):
        state = 'Sub-Sub'
    # IF(AND(B5-D4<0,P1<CriticalPressure
    if temp_diff < 0 and apress < 1056:
        state = 'Sub-Super'
    # IF(AND(B5-D4>0,P1>CriticalPressure-1)
    if temp_diff > 0 and apress > 1055:
        state = 'Super-Sub'
    if temp_diff < 0 and apress > 1055:
        state = 'Super-Super'
    Dyn_Viscosity = Form_Interp(state, visc_data)
    # -----------------------------------------------------------------------------------------------

    return [Density, Enthalpy, Dyn_Viscosity, press, Entropy, temp]
# Superheated refrigerant properties function:
#   ref: refrigerant
#   EVAPtemp: Evaporation temperature
#   CONDtemp: Condensing temperature
#   ent_DischTemp: User entered discharge temperature
def SHRefProp(ref, EVAPtemp, CONDtemp, ent_DischTemp= None, ent_Pressure = None):

    #Setting Super-heat value
    if EVAPtemp >= 0:
        SH_Sucttemp = EVAPtemp + 20
    if EVAPtemp < 0:
        SH_Sucttemp = EVAPtemp + 30
    if SH_Sucttemp > 65:
        SH_Sucttemp = 65
    CONDtemp = math.ceil(CONDtemp)

    #reading superheated vapor data file
    inFile1 = 'SH_HFC.Refprop'
    fd1 = open(inFile1, 'rb')
    data = pickle.load(fd1)

    refdata = data.loc[data['REF_ID'] == ref]

    #Suction vapor data
    EVAPdata = refdata.loc[refdata['TEMP_SATURATED'] == EVAPtemp]
    EVAPdata = EVAPdata.loc[EVAPdata['TEMP_SUPERHEATED'] == SH_Sucttemp]
    EVAPdata = EVAPdata.reset_index(drop = True)
    V_entropy = EVAPdata.loc[0] ['ENTROPY']

    # Discharge entropy value
    # 70% efficiency (R&D recommendations)
    D_entropy = V_entropy * 0.7

    # 100% efficiency (SvD version)
    D_entropy = V_entropy

    #Discharge Vapor data
    if ref == 'R-744':
        if (CONDtemp > 87 and ent_Pressure > 1060) and (ent_DischTemp is None or ent_Pressure is None):
            print(colored('Error - Need Discharge temp and pressure entry for supercritical systems', 'magenta'))
            sys.exit(1)
        # Translated R&D tool
        RefProps = CO2_RefProp(ent_DischTemp, ent_Pressure)

    if ref != 'R-744':
        CONDdata = refdata.loc[refdata['TEMP_SATURATED'] == CONDtemp]
        CONDdata = CONDdata.reset_index(drop=True)

        # Method used to find isentropic compression discharge temperature
        SH_entropy = []
        n = 0
        numRows = len(CONDdata.loc[:, 'TEMP_SATURATED'])
        for n in range(numRows):
            entry = CONDdata.loc[n]['ENTROPY']
            SH_entropy.append(entry)
        ent_diff = list(abs(SH_entropy - D_entropy))
        SH_Dischtemp_Data = list(CONDdata.loc[:, 'TEMP_SUPERHEATED'])
        if ent_DischTemp != None:
            SH_Dischtemp_Pos = SH_Dischtemp_Data.index(ent_DischTemp)
        if ent_DischTemp == None:
            SH_Dischtemp_Pos = ent_diff.index(min(ent_diff))

        # Discharge temperature
        SH_Dischtemp = CONDdata.loc[SH_Dischtemp_Pos]['TEMP_SUPERHEATED']

        # Super-heated properties
        Density = CONDdata.loc[SH_Dischtemp_Pos]['DENSITY']
        Enthalpy = CONDdata.loc[SH_Dischtemp_Pos]['ENTHALPY']
        Press = CONDdata.loc[SH_Dischtemp_Pos]['PRESSURE']
        Entropy = D_entropy
        Dyn_Viscosity = CONDdata.loc[SH_Dischtemp_Pos]['VISCOSITY']
        RefProps = [Density, Enthalpy, Dyn_Viscosity, Press, Entropy, SH_Dischtemp]
    return RefProps

def CalculateF(diameter,roughness,reynolds):
    # Starting Friction Factor
    friction = 0.01
    while 1:
        # Solve Left side of Eqn
        leftF = 1 / friction**0.5

        # Solve Right side of Eqn
        rightF = - 2 * math.log10((2.51/(reynolds * friction**0.5))+(roughness/(3.7*diameter)))

        # Change Friction Factor
        friction = friction + 0.0001
        if leftF - rightF <= 0.001:  # Check if Left = Right
            break
    return friction

# Line sizing function returns the properties [Velocity, Pressure Drop, Mass Flow, Reynolds Number] of a line size based off of the fluid properties
#   ref: refrigerant
#   LineTYP: Suction, Liquid, or Discharge Line Type
#   EVAPtemp: Evaporation temperature (°F)
#   CONDtemp: Condensing temperature (°F)
#   LIQtemp: Liquid temperature (°F)
#   LineOD: Outer diameter of line size (in)
#   PipeRUN: Length of run of the line being investigated (ft)
#   HeatLOAD: Heat load of case load/compressor capacity (BTU/h)
#   ent_MassFlow: Mass flow value entered by user (lb/h)
def LineSizing (refrigerant, LineTYP, EVAPtemp, CONDtemp, LIQtemp, LineID, PipeRUN, Material, HeatLOAD=None, ent_MassFlow=None,
                ent_DischTemp=None, ent_Pressure=None):

    #############################################################################################
    ######################### Review and update ###################################################
    NomCU = [1/4, 3/8, 1/2, 5/8, 7/8, 1.125, 1.375, 1.625, 2.125, 2.625, 3.125, 4.125, 6.125]
    NomSTL = [1/2, 3/4, 1, 1.25, 1.5, 2, 2.5, 3, 4, 6, 8]

    # Material ID
    if Material == 'CU-L':
        Nom = NomCU
        LineID = [0.19, 0.311, 0.43, 0.545, 0.785, 1.025, 1.265, 1.505, 1.985, 2.465, 2.945, 3.905, 5.845]
    if Material == 'CU-K':
        Nom = NomCU
        LineID = [0.19, 0.305, 0.402, 0.527, 0.745, 0.995, 1.245, 1.481, 1.959, 2.435, 2.907, 3.857, 5.741]
    if Material == 'CU-C194':
        Nom = NomCU[1:9]
        LineID = [0.319, 0.424, 0.531, 0.743, 0.955, 1.169, 1.381, 1.805]
    if Material == 'STL-SCH80':
        Nom = NomSTL
        LineID = [0.54, 0.75, 0.96, 1.28, 1.5, 1.94, 2.32, 2.9, 3.83, 5.76, 7.63]
    if Material == 'STL-SCH40':
        Nom = NomSTL
        LineID = [0.62, 0.82, 1.05, 1.38, 1.61, 2.07, 2.47, 3.07, 4.03, 6.07, 7.98]
    LineID = numpy.array(LineID)
    ###############################################################################################
    ###############################################################################################

    VAPProp = SatRefProp(refrigerant, EVAPtemp, "Vapor")
    LIQProp = SatRefProp(refrigerant, LIQtemp, "Liquid")
    V_enthalpy = VAPProp[1]
    L_enthalpy = LIQProp[1]

    if LineTYP == 'REF_PUMP_Suction' or LineTYP == 'REF_PUMP_Liquid':
        V_enthalpy = (V_enthalpy*.5) + (L_enthalpy*.5)
        if LineTYP == 'REF_PUMP_Suction':
            Dyn_Vis = LIQProp[2]
            Density = LIQProp[0]
        else:
            Dyn_Vis = (VAPProp[2]) * .5 + (LIQProp[2]) * .5
            Density = (VAPProp[0]) * .5 + (LIQProp[0]) * .5
    else:
        if LineTYP == 'Suction':
            Dyn_Vis = VAPProp[2]
            Density = VAPProp[0]
        if LineTYP == 'Liquid':
            Dyn_Vis = LIQProp[2]
            Density = LIQProp[0]
        if LineTYP == 'Discharge':
            # Super-heated properties
            SHProp = SHRefProp(refrigerant, EVAPtemp, CONDtemp, ent_DischTemp=ent_DischTemp, ent_Pressure=ent_Pressure)
            Density = SHProp[0]
            Dyn_Vis = SHProp[2]
        if LineTYP == 'GC_Return' and refrigerant == 'R-744':
            if ent_DischTemp is None or ent_Pressure is None:
                print(colored('Entry Error - Transcritical CO2 Gas cooler return line sizing requires gas cooler temp '
                              'and gas cooler pressure entry', 'magenta'))
                sys.exit(1)
            GC_Return_Data = CO2_RefProp(CONDtemp, ent_Pressure)
            # [Density, Enthalpy, Dyn_Viscosity, press, Entropy, temp]
            Density = GC_Return_Data[0]
            Dyn_Vis = GC_Return_Data[2]
            print([Density, Dyn_Vis])
        if LineTYP == 'GC_Return' and refrigerant != 'R-744':
            print(colored('Entry Error - GC_Return line sizing is specific to transcritical CO2', 'magenta'))
            sys.exit(1)

    # Either calculating mass flow based on heat load or taking user entry mass flow value
    if ent_MassFlow is not None:
        MassFlow = ent_MassFlow
        HeatLOAD = MassFlow * (V_enthalpy - L_enthalpy)
    if HeatLOAD is not None:
        MassFlow = HeatLOAD / (V_enthalpy - L_enthalpy)
    if HeatLOAD is None and ent_MassFlow is None:
        print(colored('Error - Need either Mass flow or Heat load for calc', 'magenta'))
        sys.exit(1)

    # CS_Area: Cross-section area of pipe in in^2
    LineID = LineID / 12
    CS_Area = (math.pi * (LineID) ** 2) / 4

    # US units
    Vel = (MassFlow / (CS_Area * Density)) / 60  # ft/min
    Re = (Density * Vel * LineID) / (Dyn_Vis*60)
    rough = 0.0000591
    #rough = 1800/1000000

    ################# SI Units #####################################
    Vel_con = Vel * 0.00508  # m/sec
    Density_con = Density * 16.0185  # kg/m^3
    LineID_con = LineID*0.0254  # meter
    PipeRUN_con = (PipeRUN / 12) * 0.3048  # meter
    #################################################################

    #######################################################################################
    ########################## Need to update for matrix logic ############################
    # Colebrook White equation used in CalculateF funct
    f = []
    i = 0
    for Re_i in Re:
        if Re_i > 2000:
            f_i = CalculateF(LineID[i]*12, rough, Re_i)
        if Re_i < 2000:
            f_i = 64/Re_i
        i = i + 1
        f = numpy.append(f, f_i)
    #######################################################################################

    # Pressure Drop Calc
    Press_Drop = (f*(PipeRUN_con/LineID_con)*((Density_con*(Vel_con**2))/2))/6894.7

    # Physical properties related to line size selected
    Pipe_Props = numpy.column_stack((Nom, Vel, Press_Drop, Re))
    PLabel = ['Nominal Pipe Size (in)', 'Velocity (m/s)', 'Pressure Drop (psi)', 'Reynolds Num']
    Pipe_Props = numpy.vstack((PLabel, Pipe_Props))

    # Physical properties related to the temperature, pressure, and fluid selected
    Fluid_Props = [MassFlow, Density, Dyn_Vis]
    FLabel = ['Mass Flow (lb/h)', 'Density (lb/ft^3)', 'Dynamic Viscosity (lb/ft-s)']
    Fluid_Props = numpy.vstack((FLabel, Fluid_Props))

    return [Pipe_Props, Fluid_Props]

#-------FUNCTION TESTING------------------------
# def LineSizing (refrigerant, LineTYP, EVAPtemp, CONDtemp, LIQtemp, LineID, PipeRUN, HeatLOAD, ent_MassFlow, ent_Dischtemp, ent_Pressure)
LineProp = LineSizing("R-744", 'GC_Return', 20, 85.8, 34, 3.9050, 50, 'CU-C194', ent_Pressure= 1094, ent_DischTemp=222, ent_MassFlow= 11775)
print(LineProp[0])

################# Data Dump ############################################################

def csv_data_dump(filename, datalink):
    file_rd = pd.read_csv(datalink)
    outFile = filename
    fw = open(outFile, 'wb')
    pickle.dump(file_rd, fw)
    fw.close()
## User entry
# datalink =
# filename =
# csv_data_dump(filename, datalink)

###############################################################################################
###############################################################################################
# wb = load_workbook(filename= r'\\Hp-g-fs2\engr\HP-ENG\MECHANICAL\SeanTayl\Projects\HPx Eng Tool\Data\RefProp.xlsx')
# refrigerants = ['R-11', 'R-12', 'R-13', 'R-23', 'R-32', 'R-115', 'R-123', 'R-124', 'R-125', 'R-500', 'R-502', 'R-503',
#                 'R-507A', 'R-717', 'R-134A', 'R-143A', 'R-152A', 'R-401A', 'R-402A', 'R-407B', 'R-407C', 'R-407F',
#                 'R-408A', 'R-409A', 'R-410A', 'R-410B', 'R-508A', 'R-513A']
###############################################################################################
###############################################################################################
#-------GUI BUILD-------------------------------

# window_size set
# root.geometry('600x500')
# frame = Frame(root)
# frame.pack(side='top', expand=True, fill='both')
#
# #Clearing window
# def Clear():
#     for widget in frame.winfo_children():
#         widget.destroy()
#     frame.pack_forget()
#
# # Main Window
# def MainWindow():
#     Clear()
#
#     #Suction Line sizing widget
#     S_Line_Button = Button(frame, text='Suction Line Sizing Tool', command=SuctionWindow)
#     S_Line_Button.grid(row=0, column=0)
#
# #Suction tool window
# def SuctionWindow():
#     suct_w = Tk()
#     suct_w = Frame(suct_w)
#     suct_w.pack(side='top', expand=True, fill='both')
#
#     # Set Up Widgets for Suction Line Sizing Tool
#     Ref_Label = Label(suct_w, text='Refrigerant')
#     Ref_Label.grid(row=0, column=0)
#
#     Ref_Entry = Entry(suct_w)
#     Ref_Entry.grid(row=0, column=1)
#
#     LineType = 'Suction'
#
#     Evap_Label = Label(suct_w, text='Evaporator Temp (°F)')
#     Evap_Label.grid(row=1, column=0)
#
#     Evap_Entry = Entry(suct_w)
#     Evap_Entry.grid(row=1, column=1)
#
#     Cond_Temp = 'NA'
#
#     LIQ_Label = Label(suct_w, text='Liquid Temp (°F)')
#     LIQ_Label.grid(row=2, column=0)
#
#     LIQ_Entry = Entry(suct_w)
#     LIQ_Entry.grid(row=2, column=1)
#
#     ID_Label = Label(suct_w, text='Line ID (in)')
#     ID_Label.grid(row=3, column=0)
#
#     ID_Entry = Entry(suct_w)
#     ID_Entry.grid(row=3, column=1)
#
#     Main_Label = Button(suct_w, text='Return to Main', command=MainWindow())
#     Main_Label.grid(row=5, column=0)
#
#     suct_w.mainloop()
#
# S_Line_Button = Button(frame, text = 'Suction Line Sizing Tool', command = SuctionWindow)
# S_Line_Button.grid(row=0, column=0)
#
# #-------GUI NOTES-------------------------------
# # Tk() - The window
# # Label(window, text = message) - Displays a text message onto window
# # widget.pack() - Puts Item onto window where there is room
# # widget.grid(row = x, column = y) - inserts Item onto window at the coordinates x,y
# # Button(window, text = 'message', ...)
# #   state - different modes of a button (ex. DISABLED, does not allow button to be interacted w/user)
# #   padx - change the width of a button
# #   pady - change the hieght of a button
# #   command - calls a function (ex. command = function name)
# #   fg - foreground color (ex. fg = 'black')
# #   bg - background color (ex. bg = 'green')
# #       *can use hex color codes
# # Entry(window, ...)
# #   width - width of entry line
# #   fg - foreground color (ex. fg = 'black')
# #   bg - background color (ex. bg = 'green')
# #   borderwidth - increase the border width
# #   entry_variable.get() - returns entry text as a string
# #   entry_variable.insert(0, "text:") - text that will appear before the user's entry
#
# #-------------------------------GUI WINDOW LOOP----------------------------------------------------------
# root.mainloop()