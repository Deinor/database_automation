import openpyxl
wb = openpyxl.load_workbook('final-confirmed_2022_6_23.xlsx')
wb.sheetnames           # The workbook's sheets' names.
sheet = wb ["CAD"]      #Get a sheet from workbook  
print(sheet)            #Debug_1

project_number = sheet["A5"].value  #Get project name
pjm_de = sheet["D5"].value
ordered = sheet["F5"].value  #Get ordered hours
used = sheet["G5"].value  #Get used hours
available = sheet["A5"].value  #Get available hours
print(project_number, pjm_de, ordered, used, available)   #Debug_2