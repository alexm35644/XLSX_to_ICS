from datetime import datetime
from ics import Calendar, Event
import os
import sys
import openpyxl 


#FILE STUFF
if getattr(sys, 'frozen', False):  # Check if running as a frozen executable
    # Running as an executable, so get the directory where the executable was run from
    executable_path = os.path.dirname(sys.executable)  # Get the directory of the executable
else:
    # Running as a script
    executable_path = os.path.dirname(os.path.abspath(__file__))  # Get the script's location

file_path = os.path.join(executable_path, "bruh.ics")
file_path_excel = os.path.join(executable_path, "View_My_Courses.xlsx")

#EXCEL FILE READING
wb = openpyxl.load_workbook(file_path_excel)
sheet = wb.active

#################################
#   ASSIGNING CALENDAR EVENTS
#################################
c = Calendar()

for row in sheet.iter_rows(min_row=4, max_row=5):
    e = Event()
    # Access data from columns B (5), H (8), K (11), L (12)
    section = row[4].value  # Column B (5tj column)
    info = row[7].value       # Column H (8th column)
    start = row[10].value      # Column K (11th column)
    end = row[11].value      # Column L (12th column)

    #Event Name is based on the course section
    e.name = section

    #Event begginging date is based on the start date and the meeting time
    e.begin = datetime.fromisoformat("2024-12-19T19:00:00-08:00")
    e.end = datetime.fromisoformat("2024-12-19T23:00:00-08:00")
    c.events.add(e)


# #CALENDAR EVENT CREATION
# c = Calendar()
# e = Event()
# e.summary = "My cool event"
# e.name = "party"
# e.description = "A meaningful description"
# e.begin = datetime.fromisoformat("2024-12-19T19:00:00-08:00")
# e.end = datetime.fromisoformat("2024-12-19T23:00:00-08:00")

# # Add the event to the calendar
# c.events.add(e)

# Try to write the calendar file

with open(file_path, "w", newline='') as f:
     f.write(c.serialize())


