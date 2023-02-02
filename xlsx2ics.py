import sys
import openpyxl as xl
import icalendar as ics
import csv
from uuid import uuid4
from datetime import datetime, timedelta
from platform import uname

DEBUG=True
FORMAT_TIME="%H:%M"
LAPSETIME=4
C_START_DATE=0
C_START_TIME=1
C_END_DATE=0
C_NAME=2
C_DISTANCE=3
C_ELEVATION=5
C_DESCRIPTION=4

def is_excel_file(filename):
    return filename.split(".")[-1] == "xlsx"

def check_args():
    success = False
    if len(sys.argv)<2:
        print("Too few command-line arguments")
    elif len(sys.argv)>2:
        print("Too many command-line arguments")
    elif not is_excel_file(sys.argv[1]):
        print("Not a xlsx file")
    else: success = True
    return success

def ics_from_xlsx(xlsxfile):
    wb = xl.load_workbook(xlsxfile)
    sh = wb.active
    cal = ics.Calendar()
    cal.add('prodid', '-//FAA Calendar//bxsoft.com//')
    cal.add('version', '1.0')
    ics_file = open(f"{xlsxfile.split('.xlsx')[0]}.ics", 'wb')

    for row in sh.iter_rows(min_row=3,values_only=True):
        if DEBUG:
            print(row)
        event = ics.Event()
        t = datetime.strptime(row[C_START_TIME], FORMAT_TIME)
        start_date = row[C_START_DATE] + timedelta(hours=t.hour, minutes=t.minute)
        end_date = start_date + timedelta(hours=LAPSETIME)
        distance = ''.join([n for n in row[C_DISTANCE] if n.isdigit()])
        elevation = int(''.join([n for n in row[C_ELEVATION] if n.isdigit()]))
        event.add('summary', f"{row[C_NAME]} - {distance}kms - {elevation:,}m.d.a")
        event.add('dtstart', start_date)
        event.add('dtend', end_date)
        event.add('description', row[C_DESCRIPTION])
        event.add('location', row[C_NAME])
        event.add('uid', uuid4().hex + '@'+ uname().node)
        event.add('dtstamp', datetime.now())
        cal.add_component(event)

    if DEBUG:
        print(cal.to_ical().decode("utf-8")) 
    ics_file.write(cal.to_ical())
    ics_file.close()

if __name__=="__main__":
    if check_args():
        ics_from_xlsx(sys.argv[1])
    else:
        sys.exit(1)

