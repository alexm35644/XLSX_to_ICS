from datetime import datetime
from icalendar import Calendar, Event
from datetime import datetime
import zoneinfo
import os
import sys
import openpyxl 
import re


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

# Time formatting function
def extractTime(text):
    # Function to convert time to 24-hour format
    def convert_to_24hr(time_str):
        # Convert 'a.m.' and 'p.m.' to 'AM' and 'PM' for compatibility
        time_str = time_str.replace("a.m.", "AM").replace("p.m.", "PM")
        return datetime.strptime(time_str, "%I:%M %p").strftime("%H:%M")

    # Extract the time range using regular expressions (considering a.m. and p.m.)
    time_range = re.search(r"(\d{1,2}:\d{2} [ap\.m]+) - (\d{1,2}:\d{2} [ap\.m]+)", text)

    if time_range:
        # Convert start and end time to 24-hour format
        start_time_24hr = convert_to_24hr(time_range.group(1).strip())
        end_time_24hr = convert_to_24hr(time_range.group(2).strip())

        # Split hour and minute into separate elements and store them in a flat list
        start_hour, start_minute = map(int, start_time_24hr.split(":"))
        end_hour, end_minute = map(int, end_time_24hr.split(":"))

        # Store each value in a flat 4x1 list
        time_list = [start_hour, start_minute, end_hour, end_minute]
        
        return time_list
    else:
        return None
    
#Location Information 
def extractLocation(text):
    # Clean the text by removing unwanted newlines or extra spaces
    cleaned_text = text.replace("\n", " ").strip()

    # Split the string by pipe symbol and get the value after the third pipe
    parts = cleaned_text.split(" | ")

    # Check if we have at least 4 parts (including the location)
    if len(parts) >= 4:
        # The location is the last part, so return it after stripping any extra spaces
        return parts[3].strip()
    else:
        return None



#################################
#   ASSIGNING CALENDAR EVENTS
#################################
cal = Calendar()

for row in sheet.iter_rows(min_row=4):
    event = Event()
    # Access data from columns B (5), H (8), K (11), L (12)
    section = row[4].value  # Column B (5th column)
    info = row[7].value       # Column H (8th column)
    start = row[10].value      # Column K (11th column)
    end = row[11].value      # Column L (12th column)

    #Event Name is based on the course section
    event.add('summary', section)

    #EVENT TIME
    Year = start.year
    Month = start.month
    Day = start.day
    time_array = extractTime(info) #Extracting class time from the info column
    startHour = time_array[0]
    startMinute = time_array[1]
    endHour = time_array[2]
    endMinute = time_array[3]
    Second = 0
    event.add('dtstart', datetime(Year, Month, Day, startHour, startMinute, Second, tzinfo=zoneinfo.ZoneInfo("America/Vancouver")))
    event.add('dtend', datetime(Year, Month, Day, endHour, endMinute, Second, tzinfo=zoneinfo.ZoneInfo("America/Vancouver")))

    #EVENT Location
    location = extractLocation(info)
    event.add('description', location)
    
    #Event Recurrence 
    Year = end.year
    Month = end.month
    Day = end.day
    Hour = 0
    Minute = 0
    Second = 0
    end = datetime(Year, Month, Day, Hour, Minute, Second, tzinfo=zoneinfo.ZoneInfo("America/Vancouver"))
    days_of_week = info.split('|')[1].strip()
    formatted_days = [day[:2].lower() for day in days_of_week.split()]
    print(formatted_days)
    print('mo')
    event.add('RRULE', {'freq': 'weekly', 'interval': '1', 'until': end, 'BYDAY':formatted_days})
    event.add("VERY", 'SLOW')

    #Add to calendar 
    cal.add_component(event)



    
#Write to .ics file
with open(file_path, "wb") as f:
     f.write(cal.to_ical())
     f.close()


