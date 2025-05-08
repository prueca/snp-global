'''
This script combines all sheets from all excel files (.xls)
within a directory into one excel file (.xlsx).

Note: All sheets are assumed to have the same headers.
Also, the first and second row of first sheets are removed
as we assume the third row contains the header.
'''

import os
import glob
import pandas as pd
from datetime import datetime
import time

start_time = datetime.now()

all_files = glob.glob(os.path.join(r".\excel", "*.xls"))
outputFile = r".\output.xlsx"

if os.path.exists(outputFile):
  os.remove(outputFile)

try:
  all_sheets = []

  with pd.ExcelWriter(outputFile, engine="xlsxwriter") as writer:
    for file in all_files:
      xl = pd.ExcelFile(file)
      first_sheet = True

      for sheet_name in xl.sheet_names:
        print(f"{file} - {sheet_name}")
        sheet = xl.parse(sheet_name, header=None)

        if (first_sheet):
          sheet.drop(index=0, axis=0, inplace=True)
          sheet.drop(index=1, axis=0, inplace=True)
          first_sheet = False
          
        all_sheets.append(sheet)

    combined = pd.concat(all_sheets, ignore_index=True)

    # remove blank rows
    print("Removing blank rows...")
    combined.dropna(how="all", inplace=True)

    # set first row as header
    combined.rename(columns=combined.iloc[0], inplace=True)

    # remove duplicates
    print("Removing duplicates...")
    combined.drop_duplicates(inplace=True)

    print("Writing excel file...")
    combined.to_excel(writer, sheet_name="Sheet1", index=False)

except Exception as err:
  print(f"An error occurred: {err}")

time.sleep(5)
end_time = datetime.now()

time_difference = end_time - start_time
minutes_duration = int(time_difference.total_seconds() / 60)
seconds_duration = int(time_difference.total_seconds() % 60)

print(f"Duration: {minutes_duration}m {seconds_duration}s")