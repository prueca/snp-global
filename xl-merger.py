'''
This script combines all sheets from all excel files (.xls)
within a directory into one excel file (.xlsx).

Note: All sheets must have the same column headers.
'''

import os
import glob
import pandas as pd

all_files = glob.glob(os.path.join(r".\excel", "*.xls"))
outputFile = r".\output.xlsx"

if os.path.exists(outputFile):
  os.remove(outputFile)

try:
  all_sheets = []

  with pd.ExcelWriter(outputFile, engine="xlsxwriter") as writer:
    for file in all_files:
      xl = pd.ExcelFile(file)

      for sheet_name in xl.sheet_names:
        sheet = xl.parse(sheet_name, header=None)
        all_sheets.append(sheet)

    combined = pd.concat(all_sheets, ignore_index=True)

    # remove blank rows
    combined.dropna(how="all", inplace=True)

    # set first row as header
    combined.rename(columns=combined.iloc[0], inplace=True)

    # remove duplicates
    combined.drop_duplicates(inplace=True)

    combined.to_excel(writer, sheet_name="Sheet1", index=False)

except Exception as err:
  print(f"An error occurred: {err}")