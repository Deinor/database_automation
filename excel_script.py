import os 
import openpyxl
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string    #Import for geting row number



def import_os():
    print(os.listdir("."))

def get_data (row, sheet):  #Get row number
    used_data = {}
    
    used_data["project_number"] = sheet ["A"+str(row)].value  #Get project name
    used_data["s_nr"] = sheet ["C" + str(row)].value #Get S-Number
    used_data["pjm_de"] = sheet["D"+ str(row)].value
    used_data["ordered"] = sheet["F"+ str(row)].value  #Get ordered hours
    used_data["used_hours"] = sheet["G"+ str(row)].value  #Get used hours
    used_data["available"] = sheet["H" + str(row)].value  #Get available hours
    return used_data
    #print("Excel script", used_data)    #Debug_1_3

import_os()
#test = get_listssnr()
#print(get_listssnr())
#test3 = get_row()
#test4 = get_data()


#print(row, "Global")
#print(list_snr, "Global")   #Debug_2_1
#print (used_data, "global")    #Debug_2_2