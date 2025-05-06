import os
import glob
import pandas as pd

all_files = glob.glob(os.path.join(".\\excel", "*.xls"))
outputFile = ".\\output.xlsx"

try:
  with pd.ExcelWriter(outputFile, mode="w") as writer:

    for file in all_files:
      # open excel file to get sheet names
      xl = pd.ExcelFile(file)

      for sheet_name in xl.sheet_names:
        # get sheet from excel file by name
        sheet = pd.read_excel(file, sheet_name=sheet_name)

        # write to output file
        sheet.to_excel(writer, sheet_name=sheet_name)

except Exception as err:
  print(f"An error occurred: {err}")