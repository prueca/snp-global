'''
This script combines all sheets from all excel files (.xls)
within a directory into one excel file (.xlsx).

Note: All sheets must have the same column headers.
'''

import os
import glob
import pandas as pd

all_files = glob.glob(os.path.join(".\\excel", "*.xls"))
outputFile = ".\\output.xlsx"

try:
  all_sheets = []

  with pd.ExcelWriter(outputFile, mode="w") as writer:
    for file in all_files:
      xl = pd.ExcelFile(file)

      for sheet_name in xl.sheet_names:
        sheet = xl.parse(sheet_name)
        all_sheets.append(sheet)

    combined = pd.concat(all_sheets, ignore_index=True)
    combined.to_excel(writer, sheet_name="Sheet1", index=False)

except Exception as err:
  print(f"An error occurred: {err}")