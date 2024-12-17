from datetime import datetime
from icalendar import Calendar, Event
import zoneinfo
import os
import sys
import openpyxl 
import re
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QFileDialog

# FILE STUFF
if getattr(sys, 'frozen', False):  # Check if running as a frozen executable
    executable_path = os.path.dirname(sys.executable)
else:
    executable_path = os.path.dirname(os.path.abspath(__file__))

file_path = os.path.join(executable_path, "schedule.ics")

# Time formatting function
def extractTime(text):
    def convert_to_24hr(time_str):
        time_str = time_str.replace("a.m.", "AM").replace("p.m.", "PM")
        return datetime.strptime(time_str, "%I:%M %p").strftime("%H:%M")

    time_range = re.search(r"(\d{1,2}:\d{2} [ap\.m]+) - (\d{1,2}:\d{2} [ap\.m]+)", text)

    if time_range:
        start_time_24hr = convert_to_24hr(time_range.group(1).strip())
        end_time_24hr = convert_to_24hr(time_range.group(2).strip())

        start_hour, start_minute = map(int, start_time_24hr.split(":"))
        end_hour, end_minute = map(int, end_time_24hr.split(":"))

        return [start_hour, start_minute, end_hour, end_minute]
    else:
        return None

# Location Information 
def extractLocation(text):
    cleaned_text = text.replace("\n", " ").strip()
    parts = cleaned_text.split(" | ")

    if len(parts) >= 4:
        return parts[3].strip()
    else:
        return None


class MyWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("XLSX Schedule to ISCS")
        self.setGeometry(100, 100, 300, 200)

        # Create a QLabel widget
        self.label = QLabel("Click to Convert File", self)

        # Create a QPushButton widget
        self.button = QPushButton("Convert", self)
        self.button.setVisible(False)  # Initially hidden
        self.button.clicked.connect(self.on_button_click)

        # Create a button to simulate the file import action
        self.import_button = QPushButton("Import File", self)
        self.import_button.clicked.connect(self.import_file)

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
        self.file_path_excel, _ = QFileDialog.getOpenFileName(self, "Select File", "", "All Files (*)", options=options)
        
        if self.file_path_excel:
            self.label.setText(f"File imported: {self.file_path_excel}")
        else:
            self.label.setText("No file selected.")
        
        # Re-enable the import button and show the second button
        self.button.setVisible(True)  # Show the second button after import
        self.import_button.setVisible(False)  # Hide import button after import
        self.import_button.setEnabled(True)  # Re-enable import button

    def on_button_click(self):
        if not self.file_path_excel:
            self.label.setText("No file selected.")
            return
        
        # Read and process the Excel file
        wb = openpyxl.load_workbook(self.file_path_excel)
        sheet = wb.active

        cal = Calendar()

        for row in sheet.iter_rows(min_row=4):
            event = Event()
            section = row[4].value
            info = row[7].value
            start = row[10].value
            end = row[11].value

            event.add('summary', section)
            time_array = extractTime(info)
            startHour, startMinute, endHour, endMinute = time_array
            Second = 0
            event.add('dtstart', datetime(start.year, start.month, start.day, startHour, startMinute, Second, tzinfo=zoneinfo.ZoneInfo("America/Vancouver")))
            event.add('dtend', datetime(end.year, end.month, end.day, endHour, endMinute, Second, tzinfo=zoneinfo.ZoneInfo("America/Vancouver")))

            location = extractLocation(info)
            event.add('description', location)

            # Recurrence (for example)
            days_of_week = info.split('|')[1].strip()
            formatted_days = [day[:2].lower() for day in days_of_week.split()]
            event.add('RRULE', {'freq': 'weekly', 'interval': '1', 'until': end, 'BYDAY':formatted_days})

            cal.add_component(event)

        # Write to .ics file
        with open(file_path, "wb") as f:
            f.write(cal.to_ical())

        # After file conversion is done
        self.label.setText("Done")
        self.button.setVisible(False)  # Hide the Convert button
        self.import_button.setVisible(False)  # Hide the Import button
        

# Run the application
app = QApplication([])
window = MyWindow()
window.show()
app.exec_()
