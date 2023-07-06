# CODE WRITTEN BY JOSHUA LEE,
# GITHUB: scaryjoshie
# EMAIL: scaryjoshie@gmail.com


# Imports
import glob
import os
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# Python Imports
from src.TableOperations import Table
from src.MiscUtils import location_to_cell_id
from src.DuplicateShift import duplicate_shift


######################################################################


# INPUT FOLDER PATH
FOLDER_PATH = "Tables\\1970 Mar Apr Right Tables"
# Gets paths of every xlsx file in specified directory
file_paths = glob.glob(f"{FOLDER_PATH}/*.xlsx")
# CONSTANTS
QUOTIENT_MAX = (
    10  # max values that 2 cells can have when divided by one another before the row is flagged
)


######################################################################


color_dict = {
    "date": "fffffb85",  # date --> yellow (#fffb85)
    "normal": "fffc9dde",  # normal anomaly --> pink (#fc9dde)
    "label": "ff71bd74",  # label --> green (#71bd74)
    "row_culprit": "ff72a5f7",  # row_culprit --> dark blue (#515be0)
    "row_small": "ff4de3e0",
    "no_decimal": "ffd088f2",
    "row": "ffbbd4fc",  # row --> light blue (#96b1ff)
}

fillers = {}
for color in color_dict.keys():
    temp = PatternFill(patternType="solid", fgColor=color_dict[color])
    fillers[color] = temp

######################################################################


# Creates directory for output
output_dir = "output18"
if os.path.exists(os.path.join("output", output_dir)):
    while True:
        overwrite = input("Dir w/ same name exists. Overwrite? (y/n): ")
        if "y" == overwrite:
            os.rmdir(os.path.join("output", output_dir))
            break
        elif "n" == overwrite:
            output_dir = input("Choose a new output dir name: ")
            break
        else:
            continue
if not os.path.exists("output"):
    os.mkdir("output")
os.mkdir(os.path.join("output", output_dir))


# Runs through every file
for file_path in file_paths:
    print(file_path)

    # Reads file into df then runs through the demerger
    raw_df = pd.read_excel(file_path, "Table_0")
    df = duplicate_shift(raw_df)

    # Creates table object and checks all values
    table = Table(df=df)
    table.CheckDates
    table.CheckNormals(QUOTIENT_MAX=QUOTIENT_MAX)
    table.CheckLabels()

    # Loads file into openpyxl
    wb = openpyxl.load_workbook(file_path)
    ws = wb["Table_0"]
    rows = dataframe_to_rows(table)

    # Fills every cell value with what it's supposed to be
    for cell in table.df_list:
        ws[location_to_cell_id(cell.location)] = cell.value

    # Fills in all colors
    for type in table.bad_cells.keys():
        for cell in table.bad_cells[type]:
            ws[location_to_cell_id(cell.location)].fill = fillers[type]

    # Names and saves file
    name = file_path.split("\\")[-1]
    wb.save(f"Output\\{output_dir}\\{name}")
