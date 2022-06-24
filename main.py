import openpyxl
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string    #Import for geting row number
wb = openpyxl.load_workbook('final-confirmed_2022_6_23.xlsx')
wb.sheetnames           # The workbook's sheets' names.
sheet = wb ["CAD"]      #Get a sheet from workbook  
print(sheet)            #Debug_1

#Get row number
xy = coordinate_from_string ("D5")
row = xy [1]
print(row)          #Debug_1_1


#Will change to Dictionarie
project_number = sheet ["A"+str(row)].value  #Get project name
pjm_de = sheet["D"+ str(row)].value
ordered = sheet["F"+ str(row)].value  #Get ordered hours
used = sheet["G"+ str(row)].value  #Get used hours
available = sheet["H" + str(row)].value  #Get available hours
print(project_number, pjm_de, ordered, used, available)   #Debug_2
