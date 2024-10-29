import argparse
import math
from datetime import datetime

import pandas as pd

from ics import Calendar, Event

# Utils

# Utils : Parse Date

def parseDateUtil(date_value):
    """Utility function to parse a date value to datetime if it's not already."""
    return date_value if isinstance(date_value, datetime) else datetime.strptime(str(date_value), '%m/%d/%Y')


# Arguments

parser = argparse.ArgumentParser(description="Convert an XLS file to ICS format.")
parser.add_argument("input_path", help="Path to the input .xls file")
parser.add_argument("output_path", help="Path to the output .ics file")
args = parser.parse_args()

# Excel

# Excel : Load
df = pd.read_excel(args.input_path, skiprows=2, header=None)

# ICS

calendar = Calendar()

# ICS : Data (from Excel)
for _, row in df.iterrows():
    # Clause Guard : Empty dates
    if pd.isna(row[1]) or pd.isna(row[2]):
        print(f"Skipping row {_} due to missing dates.")
        continue
    
    # Event

    event = Event()
    event.name = row[3]
    event.begin = parseDateUtil(row[1])
    event.end = parseDateUtil(row[2])

    # Event : Add to ICS calendar
    calendar.events.add(event)

# ICS : File : Write
with open(args.output_path, "w") as f:
    f.writelines(calendar)

# Result

print(f"ICS file created successfully at {args.output_path}")
