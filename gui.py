import PyQt5.QtWidgets as qtw
import PyQt5.QtGui as qtg
import openpyxl
from openpyxl.utils.cell import coordinate_from_string
from excel_script import get_data

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
    
#print(get_data(sheet))
#print()

class MainWindow(qtw.QWidget):
    def __init__(self):
        super().__init__()
        
        #Add Title
        self.setWindowTitle("Johny XXX")

        #Set Vertical layout
        self.setGeometry(800, 300, 600, 600)
        self.setLayout(qtw.QVBoxLayout())

        #Create lable for list SNR
        label_snr = qtw.QLabel("Select S-Number:")
       
        #Change font size of lable
        label_snr.setFont(qtg.QFont("Helvetica", 15))
        self.layout().addWidget(label_snr)

        #Create Combo box
        my_combo = qtw.QComboBox(self)

        #Add Items to combo Box
        my_combo.addItems(list_snr)

        #Put combo on screen
        self.layout().addWidget(my_combo)

        #Create a button
        button_generateOut = qtw.QPushButton("Generate e-mail", clicked = lambda: generateOut())
        self.layout().addWidget(button_generateOut)

        #Show the all
        self.show()

        #Generate out. data
        def generateOut():
            dic = get_data(sheet)
            print(dic)

app = qtw.QApplication([])
mw = MainWindow()

#Run window
app.exec_()