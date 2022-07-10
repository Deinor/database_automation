#Creating dictionary with keys and empty values. Used values needs to be epty to create first instance of table. Will be replaced by data from project after generate_data) 
used_dataFC = {"project_number": "",
            "s_nr": " ",
            "pjm_de": " ",
            "pjm_de": " ",
            "ordered": " ",
            "used_hours": " ",
            "available": " ",
            "used_lastm_hours": " "}

def get_dataFC (row_FC, sheet_FC, previous_month):  #Get row_FC number
    
    if previous_month == "January": #Used month in last month
        used_dataFC["used_lastm_hours"] = sheet_FC["L" + str(row_FC)].value 
    elif previous_month == "February":
        used_dataFC["used_lastm_hours"] = sheet_FC["M" + str(row_FC)].value  
    elif previous_month == "March":
        used_dataFC["used_lastm_hours"] = sheet_FC["N" + str(row_FC)].value 
    elif previous_month == "April":
        used_dataFC["used_lastm_hours"] = sheet_FC["O" + str(row_FC)].value 
    elif previous_month == "May":
        used_dataFC["used_lastm_hours"] = sheet_FC["P" + str(row_FC)].value 
    elif previous_month == "June":
        used_dataFC["used_lastm_hours"] = sheet_FC["Q" + str(row_FC)].value 
    elif previous_month == "July":
        used_dataFC["used_lastm_hours"] = sheet_FC["R" + str(row_FC)].value 
    elif previous_month == "August":
        used_dataFC["used_lastm_hours"] = sheet_FC["S" + str(row_FC)].value 
    elif previous_month == "September":
        used_dataFC["used_lastm_hours"] = sheet_FC["T" + str(row_FC)].value 
    elif previous_month == "October":
        used_dataFC["used_lastm_hours"] = sheet_FC["U" + str(row_FC)].value 
    elif previous_month == "November":
        used_dataFC["used_lastm_hours"] = sheet_FC["V" + str(row_FC)].value 
    elif previous_month == "December":
        used_dataFC["used_lastm_hours"] = sheet_FC["W" + str(row_FC)].value 

    used_dataFC["project_number"] = sheet_FC ["A"+str(row_FC)].value  #Get project name
    used_dataFC["s_nr"] = sheet_FC ["C" + str(row_FC)].value #Get S-Number
    used_dataFC["pjm_de"] = sheet_FC["D"+ str(row_FC)].value
    used_dataFC["ordered"] = sheet_FC["F"+ str(row_FC)].value  #Get ordered hours
    used_dataFC["used_hours"] = sheet_FC["G"+ str(row_FC)].value  #Get used hours
    used_dataFC["available"] = sheet_FC["H" + str(row_FC)].value  #Get available hours
    
    
    return used_dataFC
    #print("Excel script", used_dataFC)    #Debug_1_3

#import_os()
#test = get_listssnr()
#print(get_listssnr())
#test3 = get_row()
#test4 = get_dataFC()


#print(row_FC, "Global")
#print(list_snr, "Global")   #Debug_2_1
#print (used_dataFC, "global")    #Debug_2_2