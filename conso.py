import os
import glob
import pandas as pd

loc = input("Locate excel files: ")
# loc = r"~/Documents/excel"
loc = os.path.expanduser(loc)

all_files = glob.glob(os.path.join(f"{loc}", "*.xls*"))

all = []
output = r"./out.xlsx"
# output = os.path.join(f"{loc}", "output.xlsx")

try:
    with pd.ExcelWriter(output) as writer:
        is_first = True

        for file in all_files:
            xl = pd.read_excel(file, header=None)

            if is_first == False:
                # drop row 1-3
                xl.drop(index=0, axis=0, inplace=True)
                xl.drop(index=1, axis=0, inplace=True)
                xl.drop(index=2, axis=0, inplace=True)

            all.append(xl)
            is_first = False

        combined = pd.concat(all, ignore_index=True)
        combined.to_excel(writer, sheet_name="Sheet1", index=False, header=None)

except Exception as err:
  print(f"An error occurred: {err}")