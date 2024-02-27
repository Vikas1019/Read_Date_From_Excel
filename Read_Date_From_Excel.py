import pandas as pd
import re
from datetime import datetime

# Read the Excel file into a Pandas DataFrame
df = pd.read_excel("D:/Date_Format.xlsx")

# Define your date format regex pattern to match dates in the format dd.mm.yyyy
date_format_regex = r'\b(\d{2})\.(\d{2})\.(\d{4})\b'

# Create an empty list to store matched dates
matched_dates = []

# Iterate over each cell in the DataFrame and extract dates matching the regex pattern
for column in df.columns:
    for cell in df[column]:
        if isinstance(cell, str):  # Check if the cell contains a string
            # Find all matches of the regex pattern in the cell
            dates_in_cell = re.findall(date_format_regex, cell)
            # Convert matched dates to datetime objects
            for date_tuple in dates_in_cell:
                date_str = '.'.join(date_tuple)
                matched_dates.append(datetime.strptime(date_str, '%d.%m.%Y'))

# Print the matched dates
print("All matched dates:")
for date in matched_dates:
    print(date.strftime('%d.%m.%Y'))

# Find minimum and maximum dates
if matched_dates:
    min_date = min(matched_dates)
    max_date = max(matched_dates)
    print("\nMinimum Date:", min_date.strftime('%d.%m.%Y'))
    print("\nMinimum Year:", min_date.strftime('%Y'))
    print("\nMaximum Date:", max_date.strftime('%d.%m.%Y'))
    print("\nMaximum Year:", max_date.strftime('%Y'))
else:
    print("No dates found with that format.")
