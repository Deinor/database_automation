import openpyxl
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string    #Import for geting row number



def get_excel():
    wb = openpyxl.load_workbook('final-confirmed_2022_6_23.xlsx')
    wb.sheetnames           # The workbook's sheets' names.
    global sheet
    sheet = wb ["CAD"]      #Get a sheet from workbook  
    print(sheet)            #Debug_1
    
    # Get list of projects
    s_nr_col = sheet ["C"]
    global list_snr
    list_snr = []
    for x in range(4, len(s_nr_col)):       #S-number list
        list_snr += [s_nr_col[x].value]
    print(list_snr)                         #Debug_1_1
   
def get_row ():  #Get row number
    global row
    xy = coordinate_from_string ("C5")
    row = xy [1]
    print(row)          #Debug_1_2

def get_data ():    #Will change to Dictionarie
    global used_data 
    used_data = {}
    
    used_data["project_number"] = sheet ["A"+str(row)].value  #Get project name
    used_data["s_nr"] = sheet ["C" + str(row)].value #Get S-Number
    used_data["pjm_de"] = sheet["D"+ str(row)].value
    used_data["ordered"] = sheet["F"+ str(row)].value  #Get ordered hours
    used_data["used_hours"] = sheet["G"+ str(row)].value  #Get used hours
    used_data["available"] = sheet["H" + str(row)].value  #Get available hours
    
    print(used_data)    #Debug_1_3
    

#test = get_excel()
#test3 = get_row()
#test4 = get_data()


#print(row, "Global")
#print(list_snr, "Global")   #Debug_2_1
#print (used_data, "global")    #Debug_2_2