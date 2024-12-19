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

file_path = os.path.join(executable_path, "Calendar.ics")
file_path_excel = os.path.join(executable_path, "View_My_Courses.xlsx")

#EXCEL FILE READING
wb = openpyxl.load_workbook(file_path_excel)
sheet = wb.active

# Time formatting function
def extractTime(text): #UPDATE TO EXTRACT FROM MEETING TIMES
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

import re

def extract_locations(text):
    """
    Extracts the location from a text and returns it as a list of strings.
    """
    # The regular expression captures everything after the last pipe symbol '|', which includes the location
    location_pattern = r"\| (.*)$"  
    matches = re.findall(location_pattern, text, re.MULTILINE)

    return matches if matches else None

    
def extract_time(text):
    """
    Extracts the start and end times as datetime objects in a list from an event.
    """
    time_pattern = r"(\d{1,2}:\d{2} (a.m.|p.m.)) - (\d{1,2}:\d{2} (a.m.|p.m.))"
    matches = re.findall(time_pattern, text)

    time_ranges = []
    for match in matches:
        # Remove periods from a.m./p.m.
        start_time_str = match[0].replace(".", "").strip()
        end_time_str = match[2].replace(".", "").strip()

        start_time = datetime.strptime(start_time_str, "%I:%M %p")  # %I for hour, %M for minutes, %p for AM/PM
        end_time = datetime.strptime(end_time_str, "%I:%M %p")
        time_ranges.append([start_time, end_time])
        
    return time_ranges if time_ranges else None

def extract_date_ranges(event):
    """
    Extracts all start and end dates as datetime objects in a list from an event.
    """
    date_pattern = r"(\d{4}-\d{2}-\d{2}) - (\d{4}-\d{2}-\d{2})"
    matches = re.findall(date_pattern, event)
    
    date_ranges = []
    for match in matches:
        start_date = datetime.strptime(match[0], "%Y-%m-%d")
        end_date = datetime.strptime(match[1], "%Y-%m-%d")
        date_ranges.append([start_date, end_date])
    
    return date_ranges if date_ranges else None

def extract_days_of_week(text):
    """
    Extracts and formats the days of the week from the text.
    The days are formatted as a 2D array with the first two letters in lowercase.
    """
    days_pattern = r"\| (.*?) \|"
    matches = re.findall(days_pattern, text)

    formatted_days = []
    for match in matches:
        # Split the days by spaces, take the first two letters of each, and convert to lowercase
        formatted_days.append([day[:2].lower() for day in match.split()])

    return formatted_days if formatted_days else None




#################################
#   ASSIGNING CALENDAR EVENTS   #
#################################

def generate_calendar(sheet):
    cal = Calendar()

    for row in sheet.iter_rows(min_row=4):
        
        # Access data from columns B (5), H (8), K (11), L (12)
        section = row[4].value     # Column B (5th column)
        meeting_patterns = row[7].value        # Column H (8th column)
        start = row[10].value      # Column K (11th column)
        end = row[11].value        # Column L (12th column)

        num_events = int(meeting_patterns.count('|')/3)

        for i in range (num_events): #treats each different meeting time as a new event. 

            #Event Name is based on the course section
            event = Event()
            event.add('summary', section)
            new_dates = extract_date_ranges(meeting_patterns)
            new_times = extract_time(meeting_patterns)
            new_days = extract_days_of_week(meeting_patterns)

            #EVENT TIME
            Year = new_dates[i][0].year
            Month = new_dates[i][0].month
            Day = new_dates[i][0].day
            startHour = new_times[i][0].hour
            startMinute = new_times[i][0].minute
            endHour = new_times[i][1].hour
            endMinute = new_times[i][1].minute
            Second = 0
            event.add('dtstart', datetime(Year, Month, Day, startHour, startMinute, Second, tzinfo=zoneinfo.ZoneInfo("America/Vancouver")))
            event.add('dtend', datetime(Year, Month, Day, endHour, endMinute, Second, tzinfo=zoneinfo.ZoneInfo("America/Vancouver")))

            #EVENT Location
            location = []
            location = extract_locations(meeting_patterns)
            event.add('description', location[i])
            
            #Event Recurrence 
            Year = new_dates[i][1].year
            Month = new_dates[i][1].month
            Day = new_dates[i][1].day
            Hour = 0
            Minute = 0
            Second = 0
            end = datetime(Year, Month, Day, Hour, Minute, Second, tzinfo=zoneinfo.ZoneInfo("America/Vancouver"))
            days_of_week = meeting_patterns.split('|')[1].strip()
            formatted_days = [day[:2].lower() for day in days_of_week.split()]
            event.add('RRULE', {'freq': 'weekly', 'interval': '1', 'until': end, 'BYDAY':formatted_days})
            event.add("VERY", 'SLOW')

            #Add to calendar 
            cal.add_component(event)
    return cal



cal = generate_calendar(sheet)
#Write to .ics file
with open(file_path, "wb") as f:
     f.write(cal.to_ical())
     f.close()


