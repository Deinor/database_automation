from operator import index
import PyQt5.QtWidgets as qtw
import PyQt5.QtGui as qtg
import openpyxl
from openpyxl.utils.cell import coordinate_from_string
from excel_script import get_data, used_data

wb = openpyxl.load_workbook('final-confirmed_2022_6_23.xlsx')
wb.sheetnames           # The workbook's sheets' names.
sheet = wb ["CAD"]      #Get a sheet from workbook  
#print(sheet)            #Debug_1
    
# Get list of projects
s_nr_col = sheet ["C"]
list_snr = []
for x in range(4, len(s_nr_col)):       #S-number list
    list_snr += [s_nr_col[x].value]

#print(list_snr)                         #Debug_1_1

#Load empty used data dictionary
dic = used_data
    
class MainWindow(qtw.QWidget):
    def __init__(self):
        super().__init__()
        
        #Add Title
        self.setWindowTitle("Johny XXX")

        #Creating layouts
        self.setGeometry(600, 300, 800, 600)
        outer_layout = qtw.QVBoxLayout()
        top_layout = qtw.QFormLayout()
        body_layout = qtw.QVBoxLayout()

        #Nesting layouts
        outer_layout.addLayout(top_layout)
        outer_layout.addLayout(body_layout)

        #Set the window´s main layout
        self.setLayout(outer_layout)

        #Create lable for list SNR
        label_snr = qtw.QLabel("Select S-Number:")
       
        #Change font, text size of lable
        label_snr.setFont(qtg.QFont("Helvetica", 15))
        top_layout.addWidget(label_snr)

        #Create Combo box
        my_combo = qtw.QComboBox(self)

        #Add Items to combo Box
        my_combo.addItems(list_snr)

        #Change font, text size of lable
        my_combo.setFont(qtg.QFont("Helvetica", 10))

        #Put combo on screen
        top_layout.addWidget(my_combo)

        #Conect signals to the metods - get index of list
        def comboBox_activated(index):
            global list_snrIndex
            list_snrIndex = index
            generate_data()
            
        my_combo.activated.connect(comboBox_activated)

        #Create table 
        table1 = qtw.QTableWidget(self)
        table1.setRowCount(1)
        table1.setColumnCount(6)
        table1.setHorizontalHeaderLabels(["Project number", "S-Number", "Project leader DE", "Ordered hours", "Used hours", "Available hours"])
        table1.setVerticalHeaderLabels(["", "", "", "", "", ""])

        #Fill table with dictionary values andd keys 
        table1.setItem(0, 0, qtw.QTableWidgetItem(str(dic["project_number"])))
        table1.setItem(0, 1, qtw.QTableWidgetItem(str(dic["s_nr"])))
        table1.setItem(0, 2, qtw.QTableWidgetItem(str(dic["pjm_de"])))
        table1.setItem(0, 3, qtw.QTableWidgetItem(str(dic["ordered"])))
        table1.setItem(0, 4, qtw.QTableWidgetItem(str(dic["used_hours"])))
        table1.setItem(0, 5, qtw.QTableWidgetItem(str(dic["available"])))
        
        #Put table on screen
        body_layout.addWidget(table1)
        
        #Create a button
        button_generateOut = qtw.QPushButton("Generate e-mail", clicked = lambda: generate_email())
        body_layout.addWidget(button_generateOut)

        #Generate out. data
        def generate_data():

            cell_address = "C" + str(list_snrIndex + 5)
            #print(cell_address) #Debug 2_2

            #Get selected row number

            xy = coordinate_from_string (cell_address)
            row = xy [1]
            #print(row)          #Debug 2_3
        
            dic = get_data(row, sheet)
            #print(dic) #Debug 2_4

            #Update table dictionary values each time project is selected in combo box
            table1.setItem(0, 0, qtw.QTableWidgetItem(str(dic["project_number"])))
            table1.setItem(0, 1, qtw.QTableWidgetItem(str(dic["used_hours"])))
            table1.setItem(0, 1, qtw.QTableWidgetItem(str(dic["s_nr"])))
            table1.setItem(0, 2, qtw.QTableWidgetItem(str(dic["pjm_de"])))
            table1.setItem(0, 3, qtw.QTableWidgetItem(str(dic["ordered"])))
            table1.setItem(0, 4, qtw.QTableWidgetItem(str(dic["used_hours"])))
            table1.setItem(0, 5, qtw.QTableWidgetItem(str(dic["available"])))

         #Generate email
        def generate_email():
            print("E-mail generated!")

        #Show the all
        self.show()

app = qtw.QApplication([])
mw = MainWindow()

#Run window
app.exec_()