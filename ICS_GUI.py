from datetime import datetime
from ics import Calendar, Event
import os
import sys

from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QVBoxLayout

if getattr(sys, 'frozen', False):  # Check if running as a frozen executable
    # Running as an executable, so get the directory where the executable was run from
    executable_path = os.path.dirname(sys.executable)  # Get the directory of the executable
else:
    # Running as a script
    executable_path = os.path.dirname(os.path.abspath(__file__))  # Get the script's location

# Now use the executable_path to save the ICS file in the correct directory
file_path = os.path.join(executable_path, "calendar_event.ics")


class MyWindow(QWidget):
    def __init__(self):
        super().__init__()

        # Set the window title and size
        self.setWindowTitle("Simple PyQt5 GUI")
        self.setGeometry(100, 100, 300, 200)

        # Create a QLabel widget
        self.label = QLabel("Hello, PyQt5!", self)

        # Create a QPushButton widget
        self.button = QPushButton("Click Me", self)
        self.button.clicked.connect(self.on_button_click)

        # Set the layout of the widgets
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.button)
        self.setLayout(layout)

    def on_button_click(self):
        # Create an event
        c = Calendar()
        e = Event()
        e.summary = "My cool event"
        e.name = "party"
        e.description = "A meaningful description"
        e.begin = datetime.fromisoformat("2024-12-19T19:00:00-08:00")
        e.end = datetime.fromisoformat("2024-12-19T23:00:00-08:00")

        # Add the event to the calendar
        c.events.add(e)

        # Try to write the calendar file
        try:
            with open(file_path, "w", newline='') as f:
                f.write(c.serialize())
            self.label.setText(f"Calendar File Created at {file_path}")
        except Exception as e:
            self.label.setText(f"Error creating file: {e}")
            print(f"Error: {e}")


# Run the application
app = QApplication([])
window = MyWindow()
window.show()
app.exec_()
