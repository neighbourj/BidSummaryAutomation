# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import openpyxl # Importing all the python packages 
import pandas as pd
import tkinter
from tkinter import *
from tkinter import ttk
from tkinter import filedialog

def GetFile(): # program to load the QE file
    import openpyxl
    import pandas as pd
    
    """
    global data_cost
    global data_labor
    global data_length
    """
    
    root = Tk() # GUI application for selecting excel file
    root.update()
    root.fileName = filedialog.askopenfilename(filetypes = (("QE File","*.xls"),("All files", "*.*"))) # opens file explorer
    file = root.fileName
    root.destroy()
    
    #file = input("Paste the file path to the QE Excel sheet: ")
    
    xl=pd.ExcelFile(file) # saves excel file as variable
    Detail = xl.parse('Detail') # grabs the column of QE that has the Spec 
    mylist = Detail['Unnamed: 0'].tolist() # turns above to list
    Length = len(mylist) # grabs length of list of specs
    Specs = mylist[8:Length] # cuts off unimportant part of list, the header of the column and above
    
    unique_spec = [] # grabs unique spec names from list
    for i in Specs:
        if i not in unique_spec and str(i) != 'nan': # need to check this to remove nan
            unique_spec.append(i)
    
    descrip_list = Detail['M&E Contractors'].tolist() # looks at description of each component, will be used below to determine which rows are for pipe
    descrip = descrip_list[8:Length-1]
    str_descrip = list(map(str, descrip))
    matl_group = ["Sch 40 Blk Steel Pipe", "Type L Hard Copper Tube", "Type K Hard Copper Tube", "Sch 40 ERW Blk Stl Pipe", "Sch 80 Blk Steel Pipe", "Sch 40 PVC BE Pipe", "Sch 80 PVC BE Pipe", "ACR Hard Copper Tube", "Sch 80 ERW Blk Stl Pipe"]
    count_decrip = 8
    index_decrip = [] 


    for i in str_descrip: # looks through description to determine which rows are for pipe and save those rows as index
        for m in matl_group: 
            if m in i:
                row_pipe_length = descrip.index(i)
                index_decrip.append(count_decrip)

        count_decrip = count_decrip + 1
        
              
    pipe_length= Detail.loc[index_decrip, "Unnamed: 1"] # grabs length of pipe column
    pipe_spec= Detail.loc[index_decrip, "Unnamed: 0"] # grabs pipe spec column

    data_length = {}
    for key, val in zip(pipe_spec,pipe_length): # combines the length of each pipe in each spec to a dictonary that has the specs as the key values
        data_length[key]= data_length.get(key, 0) + val
        
        
    labor_list = Detail['Unnamed: 14'].tolist() # grabs column for labor
    labor = labor_list[8:Length-1]
    int_labor = list(map(float, labor)) # turns labor data into list of float data


    labor_spec = mylist[8:Length-1]
   
    data_labor = {}
    for key, val in zip(labor_spec,int_labor): # sorts labor by spec key and combines the sum to dictionary 
        data_labor[key]= data_labor.get(key, 0) + val    
        
    
    
    cost_list = Detail['Unnamed: 9'].tolist() # grabs cost column
    cost = cost_list[8:Length-1]
    int_cost = list(map(float, cost))

   
    data_cost = {}
    for key, val in zip(labor_spec,int_cost): # combines labor data by spec in a dictionary 
        data_cost[key]= data_cost.get(key, 0) + val     

    big_dict = {}
    for i in unique_spec:  # combines the dictionaries for cost,  labor, and pipe length and combines in to one dictionary 
        big_dict[i] = {}
        for x in data_cost:
            big_dict[i]["cost"] = round(data_cost[i])
        for y in data_labor:
            big_dict[i]["labor"] = round(data_labor[i])
        for z in data_length:
            big_dict[i]["pipe length"] = round(data_length[i])
        
            
    flange_cost = 0
    index_flange = []
    count_flange = 8
    flange_total = 0
    
    flange_raw = Detail['M&E Contractors'].tolist() # grabs data column for flanges and selects and indexes the rows that are for flanges
    flange = flange_raw[8:Length-1]
    flange_descrip = list(map(str, flange))
    flange_type = ["PVC Blind Flange", "PVC Solv Weld Flange", "150# Cast Copper Flange", "Vic Style 741 150# Grv Flange", "CS RF Blind Flange", "CS RF Slip-On Flange", "CS RF Weld Neck Flange"]
            
    for i in flange_descrip: # creates index list for flanges
        for m in flange_type: 
            if m in i:
                row_flange = flange.index(i)
                index_flange.append(count_flange)

        count_flange = count_flange + 1     
    # sums up the total flange cost for every spec
    flange_cost= Detail.loc[index_flange, "Unnamed: 9"].tolist() # This part of the code is unnecessarily long, was getting the right output but my check was off. Fuck it tho it works
    for i in flange_cost:
        flange_total = flange_total + i

            
    #print(big_dict)
    #print(unique_spec)
    
    x= (big_dict, flange_total)
    
    return x

#GetFile()


def PasteData(x):

    from openpyxl import load_workbook
    from openpyxl import workbook
    import os

    big_dict, flange_total = x # splits tuple of returns from get file function and splits in to variables for this function

    root = Tk()
    root.update()
    root.fileName = filedialog.askopenfilename(filetypes = (("Bid Summary","*"),("All files", "*.*"))) # file explorer for destination bid summary
    fileinput = root.fileName
    root.destroy()
    
    
    #fileinput = input("Paste the file path to the Bid Summary here: ")
    file2= fileinput.replace("file:///","") # parses off part of destination file string so it can be opened properly
    
    #big_dict = {'CHW4': {'cost':53154,'labor':3378,'pipe length':2644},'CHW2': {'cost':5310,'labor':337,'pipe length':264},'HW4': {'cost':154,'labor':78,'pipe length':44}}
    
    
    
    wb = load_workbook(file2) # loads the excel file and saves the necessary mech sheet as variable sheet
    sheet = wb['MechTotal']
    
    #cost_key_raw = []
     
    CHW_labor = 0 # initializes variables for what will be posted where in bid summary
    CHW_cost = 0
    CHW_length = 0
    HW_labor = 0
    HW_cost = 0
    HW_length = 0
    CW_labor = 0
    CW_cost = 0
    CW_length = 0 
    Cond_labor = 0
    Cond_cost = 0
    Cond_length = 0    
    steam_labor = 0
    steam_cost = 0
    steam_length = 0  
    natgas_labor = 0
    natgas_cost = 0
    natgas_length = 0 
    compgas_labor = 0
    compgas_cost = 0
    compgas_length = 0 
    steamcond_labor = 0
    steamcond_cost = 0
    steamcond_length = 0 
    refrig_labor = 0
    refrig_cost = 0
    refrig_length = 0 
    
    loop_dict = [] # intialize variable that keeps track of what specs are relevant in this job for the pasting to bid summary loop later in code
    
    for key in big_dict: # loop that sums up each part of spec, for ex. CHW 1 and CHW 3 to just CHW
        if "CHW" in key: 
            CHW_labor = CHW_labor+ big_dict[key]["labor"]
            CHW_cost = CHW_cost+ big_dict[key]["cost"]
            CHW_length = CHW_length + big_dict[key]["pipe length"]
            if "CHW" not in loop_dict:
                loop_dict.append("CHW")
        if "HW" in key and "CHW" not in key:
            HW_labor = HW_labor+ big_dict[key]["labor"]
            HW_cost = HW_cost+ big_dict[key]["cost"]
            HW_length = HW_length + big_dict[key]["pipe length"]
            if "HW" not in loop_dict:
                loop_dict.append("HW")

        if "CWS" in key:
            CW_labor = CW_labor+ big_dict[key]["labor"]
            CW_cost = CW_cost+ big_dict[key]["cost"]
            CW_length = CW_length + big_dict[key]["pipe length"]  
            if "CW" not in loop_dict:
                loop_dict.append("CW")
        if "COND1" in key:
            Cond_labor = Cond_labor+ big_dict[key]["labor"]
            Cond_cost = Cond_cost+ big_dict[key]["cost"]
            Cond_length = Cond_length + big_dict[key]["pipe length"] 
            if "Condw" not in loop_dict:
                loop_dict.append("Condw")
        if "COND2" in key:
            Cond_labor = Cond_labor+ big_dict[key]["labor"]
            Cond_cost = Cond_cost+ big_dict[key]["cost"]
            Cond_length = Cond_length + big_dict[key]["pipe length"] 
            if "Condw" not in loop_dict:
                loop_dict.append("Condw")
        if "CONDRTN" in key:
            steamcond_labor = steamcond_labor+ big_dict[key]["labor"]
            steamcond_cost = steamcond_cost+ big_dict[key]["cost"]
            steamcond_length = steamcond_length + big_dict[key]["pipe length"]
            if "Condrtn" not in loop_dict:
                loop_dict.append("Condrtn")
        if "STEAM" in key:
            steam_labor = steam_labor+ big_dict[key]["labor"]
            steam_cost = steam_cost+ big_dict[key]["cost"]
            steam_length = steam_length + big_dict[key]["pipe length"]
            if "Steam" not in loop_dict:
                loop_dict.append("Steam")
        if "REFRIG" in key:
            refrig_labor = refrig_labor+ big_dict[key]["labor"]
            refrig_cost = refrig_cost+ big_dict[key]["cost"]
            refrig_length = refrig_length + big_dict[key]["pipe length"]
            if "refrig" not in loop_dict:
                loop_dict.append("refrig")
        if "GAS" in key:
            natgas_labor = natgas_labor+ big_dict[key]["labor"]
            natgas_cost = natgas_cost+ big_dict[key]["cost"]
            natgas_length = natgas_length + big_dict[key]["pipe length"]   
            if "Gas" not in loop_dict:
                loop_dict.append("Gas")
        if "AIR" in key:
            compgas_labor = compgas_labor+ big_dict[key]["labor"]
            compgas_cost = compgas_cost+ big_dict[key]["cost"]
            compgas_length = compgas_length + big_dict[key]["pipe length"] 
            if "Air" not in loop_dict:
                loop_dict.append("Air")
    #print(loop_dict)
    for i in loop_dict: # pastes spec sums for each category into their relevant place in bid summary, does not work if anything in bid summary was moved
        if "CHW" in i:
            sheet["D14"] = CHW_cost
            sheet["E14"] = CHW_labor
            sheet["F14"] = CHW_length
            sheet["I14"] = "CHW"
        if "HW" in i and "CHW" not in i:
            sheet["D15"] = HW_cost
            sheet["E15"] = HW_labor
            sheet["F15"] = HW_length
            sheet["I15"] = "HW"
        if "CW" in i:
            sheet["D16"] = CW_cost
            sheet["E16"] = CW_labor
            sheet["F16"] = CW_length
            sheet["I16"] = "CW"
        if "Condw" in i:
            sheet["D17"] = Cond_cost
            sheet["E17"] = Cond_labor
            sheet["F17"] = Cond_length
            sheet["I17"] = "Cond"
        if "refrig" in i:
            sheet["D18"] = refrig_cost
            sheet["E18"] = refrig_labor
            sheet["F18"] = refrig_length
            sheet["I18"] = "ref"
        if "Steam" in i:
            sheet["D19"] = steam_cost
            sheet["E19"] = steam_labor
            sheet["F19"] = steam_length
            sheet["I19"] = "stm"
        if "Condrtn" in i:
            sheet["D20"] = steamcond_cost
            sheet["E20"] = steamcond_labor
            sheet["F20"] = steamcond_length
            sheet["I20"] = "stmC"
            
    sheet["G23"]= flange_total    # pastes flange cost into relevant cell    
        
    
    wb.save(file2) # saves the excel bid summary
    wb.close()

#PasteData()

def main(): # function that runs both of the sub functions
    PasteData(GetFile())

main()