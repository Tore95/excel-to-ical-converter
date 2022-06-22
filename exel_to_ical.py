from cmath import nan
from datetime import datetime, timedelta, tzinfo
import os
from pathlib import Path
from sre_compile import isstring
from numpy import NaN
from icalendar import Calendar, Event, vCalAddress, vText
import pandas as pd
import pytz

excel_data = pd.read_excel(
    io='calendario.xlsx',
    skiprows=1,
    usecols=['DATA', 'ORARIO','ORE','ARGOMENTO']
)

course1 = 'A.1 LOGICHE DI PROGRAMMAZIONE'
course2 = 'E.1 TESTING E DEBUGGING DEL SITO '
course3 = ''

lesson_filtred = excel_data[(excel_data['ARGOMENTO'] == course1) | (excel_data['ARGOMENTO'] == course2)]
lesson_touples = [tuple(els) for els in lesson_filtred.to_numpy().tolist()]

cal = Calendar()

for tu in lesson_touples:
    time_string = tu[1][0:5].split('.')
    start_time = tu[0]
    start_time = start_time + timedelta(hours=int(time_string[0]), minutes=int(time_string[1]))
    end_time = start_time + timedelta(hours=int(tu[2][0]))
    title = tu[3]

    event = Event()
    event.add('summary', title)
    event.add('dtstart', start_time)
    event.add('dtend', end_time)

    cal.add_component(event)

dire = str(Path(__file__).parent.parent) + '/'
print("ics file generated at ", dire)
f = open(os.path.join(dire, "lessons.ics"), 'wb')
f.write(cal.to_ical())
f.close()


