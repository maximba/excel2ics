import sys
import icalendar as ics
import csv
from uuid import uuid4
from platform import uname
from datetime import datetime, timedelta

DEBUG=True
FORMAT_DATETIME="%Y-%m-%d %H:%M"
LAPSETIME=4
CSV_START_DATE=0
CSV_START_TIME=1
CSV_END_DATE=0
CSV_NAME=2
CSV_DURATION=3
CSV_ELEVATION=5
CSV_DESCRIPTION=4

def is_csv_file(filename):
    return filename.split(".")[-1] == "csv"

def check_args():
    success = False
    if len(sys.argv)<2:
        print("Too few command-line arguments")
    elif len(sys.argv)>2:
        print("Too many command-line arguments")
    elif not is_csv_file(sys.argv[1]):
        print("Not a csv file")
    else: success = True
    return success

def ics_from_csv(csvfile):
    cal = ics.Calendar()
    cal.add('prodid', '-//FAA Calendar//bxsoft.com//')
    cal.add('version', '1.0')

    csv_file = open(csvfile, 'r')
    rd = csv.reader(csv_file, quoting=csv.QUOTE_ALL)

    ics_file = open(f"{csvfile.split('.')[0]}.ics", 'wb')

    for row in rd:
        event = ics.Event()
        start_date = row[CSV_START_DATE].split(" ")[0] + " " + row[CSV_START_TIME]
        dtstart = datetime.strptime(start_date, FORMAT_DATETIME)
        event.add('summary', row[CSV_NAME])
        event.add('dtstart', dtstart)
        event.add('dtend', dtstart + timedelta(hours=LAPSETIME))
        event.add('description', row[CSV_DESCRIPTION])
        event.add('location', row[CSV_NAME])
        event.add('uid', uuid4().hex + '@' + uname().node)
        event.add('dtstamp', datetime.now())
        cal.add_component(event)

    if DEBUG:
        print(cal.to_ical().decode("utf-8")) 
    csv_file.close()
    ics_file.write(cal.to_ical())
    ics_file.close()

if __name__=="__main__":
    if check_args():
        ics_from_csv(sys.argv[1])
    else:
        sys.exit(1)

