from datetime import datetime
from icalendar import Calendar, Event
from datetime import datetime
import zoneinfo
import os
import sys
import openpyxl 
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
    Handles the special case of "(Alternate weeks)".
    """
    # Updated regex to capture the days, ignoring "(Alternate weeks)" or other similar text
    days_pattern = r"\| (.*?) \|"
    matches = re.findall(days_pattern, text)

    formatted_days = []
    for match in matches:
        # Remove the "(Alternate weeks)" part, if present
        clean_match = re.sub(r"\(.*\)", "", match).strip()
        
        # Split the days by spaces, take the first two letters of each, and convert to lowercase
        formatted_days.append([day[:2].lower() for day in clean_match.split()])

    return formatted_days if formatted_days else None

def check_alternate_weeks(text):
    """
    Checks if 'Alternate weeks' is present in each event segment defined by every third '|'.
    Returns a list where each entry is 2 if 'Alternate weeks' is found, otherwise 1.
    """
    # Split the text into event segments using 3 '|' as a separator
    events = re.findall(r'(.*?\|.*?\|.*?\|)', text, re.DOTALL)

    # Check each segment for 'Alternate weeks'
    result = [2 if "(Alternate weeks)" in event else 1 for event in events]
    return result

def find_first_occurrence(sheet, column, search_value):
    """
    Finds the first row where search_value appears in a specified column.

    :param sheet: The worksheet object from openpyxl
    :param column: Column letter to search (e.g., 'A')
    :param search_value: Value to search for
    :return: Row number if found, None otherwise
    """
    for row in sheet.iter_rows(min_col=sheet[column][0].column, max_col=sheet[column][0].column):
        for cell in row:
            if cell.value == search_value:
                return cell.row
    return None




#################################
#   ASSIGNING CALENDAR EVENTS   #
#################################

def generate_calendar(sheet, term_start):
    cal = Calendar()

    row_start = find_first_occurrence(sheet, 'H', 'Meeting Patterns')

    for row in sheet.iter_rows(min_row=(row_start+1)):
        
        
        # Access data from columns B (5), H (8), K (11), L (12)
        section = row[4].value     # Column B (5th column)
        meeting_patterns = row[7].value        # Column H (8th column)
        start = row[10].value      # Column K (11th column)
        end = row[11].value        # Column L (12th column)

        if(meeting_patterns.count('|')==0): # Stops the loop. This is if the schedule contains waitlists or dropped courses. 
            break

        num_events = int(meeting_patterns.count('|')/3)

        for i in range (num_events): #treats each different meeting time as a new event. 

            #Event Name is based on the course section
            event = Event()
            event.add('summary', section)
            new_dates = extract_date_ranges(meeting_patterns)
            new_times = extract_time(meeting_patterns)
            new_days = extract_days_of_week(meeting_patterns)
            new_weeks = check_alternate_weeks(meeting_patterns)
    

            if new_dates[i][0].year < term_start:
                continue

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
            interval = new_weeks[i]
            print(section)
            print(interval)
            Yearend = new_dates[i][1].year
            Monthend = new_dates[i][1].month
            Dayend = new_dates[i][1].day
            Hourend = 0
            Minuteend = 0
            Secondend = 0
            end = datetime(Yearend, Monthend, Dayend, Hourend, Minuteend, Secondend, tzinfo=zoneinfo.ZoneInfo("America/Vancouver"))
            event.add('RRULE', {'freq': 'weekly', 'interval': interval, 'until': end, 'BYDAY':new_days[i]})
            event.add("VERY", 'SLOW')

            #Add to calendar 
            cal.add_component(event)
    return cal

