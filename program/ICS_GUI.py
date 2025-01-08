from os import path
import sys
import openpyxl 
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QFileDialog, QTextEdit
from PyQt5.QtCore import Qt
from ics_generator import generate_calendar

# FILE STUFF
if getattr(sys, 'frozen', False):  # Check if running as a frozen executable
    executable_path = os.path.dirname(sys.executable)
else:
    executable_path = path.dirname(path.abspath(__file__))

file_path = path.join(executable_path, "schedule.ics")



class MyWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("XLSX Schedule to ISCS")
        self.setGeometry(400, 200, 600, 600)

        # Create a QLabel widget
        self.label = QLabel("Click to Import File", self)
        self.label.setStyleSheet("font-size: 18px;")
        self.label.setAlignment(Qt.AlignCenter)
        self.label.setTextInteractionFlags(Qt.TextSelectableByMouse)

        # Create a QPushButton widget
        self.button = QPushButton("Click to Convert", self)
        self.button.setVisible(False)  # Initially hidden
        self.button.clicked.connect(self.File_Conversion)
        self.button.setStyleSheet("background-color: #4CAF50; color: white; font-size: 18px; padding: 10px 20px;")

        # Create a button to simulate the file import action
        self.import_button = QPushButton("Import File", self)
        self.import_button.clicked.connect(self.import_file)
        self.import_button.setStyleSheet("background-color: #47A6F9; color: white; font-size: 18px; padding: 10px 20px;")





        # Set the layout of the widgets
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.import_button)
        layout.addWidget(self.button)
        self.setLayout(layout)

        self.file_path_excel = None

    def import_file(self):
        self.label.setText("Importing file...")
        self.import_button.setEnabled(False)  # Disable import button during import
        
        # Open a file dialog to select a file
        options = QFileDialog.Options()
        self.file_path_excel, _ = QFileDialog.getOpenFileName(self, "Select File", "", "Excel Files (*.xlsx);;All Files (*)", options=options) # only allow .xlsx files
        
        if self.file_path_excel:
            self.label.setText(f"File imported: {self.file_path_excel}")
        else:
            self.label.setText("No file selected.")
        
        # Re-enable the import button and show the second button
        self.button.setVisible(True)  # Show the second button after import
        self.import_button.setVisible(False)  # Hide import button after import
        self.import_button.setEnabled(True)  # Re-enable import button

    def File_Conversion(self):
        if not self.file_path_excel:
            self.label.setText("No file selected.")
            return
        
        # Read and process the Excel file
        wb = openpyxl.load_workbook(self.file_path_excel)
        sheet = wb.active

        #Custom ICS logic 
        cal = generate_calendar(sheet, 2025)

        #Write to .ics file
        with open(file_path, "wb") as f:
            f.write(cal.to_ical())
            f.close()
        

        # After file conversion is done
        self.label.setText("Done - File is located at "+executable_path+"/schedule.ics")
        self.button.setVisible(False)  # Hide the Convert button
        self.import_button.setVisible(False)  # Hide the Import button
        

# Run the application
app = QApplication([])
window = MyWindow()
window.show()
app.exec_()
