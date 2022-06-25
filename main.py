import openpyxl
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string    #Import for geting row number
wb = openpyxl.load_workbook('final-confirmed_2022_6_23.xlsx')
wb.sheetnames           # The workbook's sheets' names.
sheet = wb ["CAD"]      #Get a sheet from workbook  
print(sheet)            #Debug_1

# Get list of projects
s_nr_col = sheet ["C"]
list_snr = []
for x in range(4, len(s_nr_col)):       #S-number list
    list_snr += [s_nr_col[x].value]
print(list_snr)                         #Debug_1_1

#Get row number
xy = coordinate_from_string ("C5")
row = xy [1]
print(row)          #Debug_1_2


#Will change to Dictionarie
project_number = sheet ["A"+str(row)].value  #Get project name
s_nr = sheet ["C" + str(row)]           #Get S-Number
pjm_de = sheet["D"+ str(row)].value
ordered = sheet["F"+ str(row)].value  #Get ordered hours
used = sheet["G"+ str(row)].value  #Get used hours
available = sheet["H" + str(row)].value  #Get available hours
print(project_number, pjm_de, ordered, used, available)   #Debug_2