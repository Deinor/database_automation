import PyQt5.QtWidgets as qtw
import PyQt5.QtGui as qtg
import main

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


        #Create Combo box
        my_combo = qtw.QComboBox(self)

        #Add Items to combo Box
        my_combo.addItem("One")
        my_combo.addItem("Two")
        my_combo.addItem("Three")

        #Put combo on screen
        self.layout().addWidget(my_combo)

        #Create a button
        my_button = qtw.QPushButton("Press me", clicked = lambda: press_it())
        self.layout().addWidget(my_button)



        #Show the all
        self.show()

        def press_it():
            #Add name to label
            my_label.setText(f"Hello{my_combo.currentText()}")


app = qtw.QApplication([])
mw = MainWindow()

#Run window
app.exec_()