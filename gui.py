import PyQt5.QtWidgets as qtw
import PyQt5.QtGui as qtg
from excel_script import get_excel

class MainWindow(qtw.QWidget):
    def __init__(self):
        super().__init__()
        
        #Add Title
        self.setWindowTitle("Test 1")

        #Set Vertical layout
        #self.setGeometry(800, 300, 600, 600)
        self.setLayout(qtw.QVBoxLayout())

        #Create lable
        my_label = qtw.QLabel("Select S-Number:")
       
        #Change font size of lable
        my_label.setFont(qtg.QFont("Helvetica", 15))

        self.layout().addWidget(my_label)

        #Create a button
        my_button1 = qtw.QPushButton("Load list snr", clicked = lambda: press_it())
        self.layout().addWidget(my_button1)

        #Create Combo box
        get_excel()
        my_combo = qtw.QComboBox(self)

        #Add Items to combo Box
        my_combo.addItems("1", "2")

        #Put combo on screen
        self.layout().addWidget(my_combo)

        #Create a button
        my_button2 = qtw.QPushButton("Press me", clicked = lambda: press_it2())
        self.layout().addWidget(my_button2)



        #Show the all
        self.show()

        def press_it():
            #Add name to label
            get_excel()
            #my_label.setText(f"Hello{my_combo.currentText()}")

        def press_it2():
            #Add name to label
            my_label.setText(f"Hello{my_combo.currentText()}")


app = qtw.QApplication([])
mw = MainWindow()

#Run window
app.exec_()