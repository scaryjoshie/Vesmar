# CODE WRITTEN BY JOSHUA LEE,
# GITHUB: scaryjoshie
# EMAIL: scaryjoshie@gmail.com


# Imports
import time
import glob
import os
import shutil
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from multiprocessing import Pool, Process

# Python Imports
from src.TableOperations import Table
from src.MiscUtils import location_to_cell_id
from src.DuplicateShift import duplicate_shift


######################################################################
INPUT_DIR = os.path.join("input")
OUTPUT_DIR = "output18"
QUOTIENT_MAX = (
    10  # max values that 2 cells can have when divided by one another before the row is flagged
)
######################################################################

# gets all files from INPUT_DIR
file_paths = glob.glob(f"{INPUT_DIR}/*.xlsx")

# color mapping
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


# Creates directory for output
# if os.path.exists(os.path.join("output", OUTPUT_DIR)):
#     while True:
#         overwrite = input("Dir w/ same name exists. Overwrite? (y/n): ")
#         if "y" == overwrite:
#             os.rmdir(os.path.join("output", OUTPUT_DIR))
#             break
#         elif "n" == overwrite:
#             OUTPUT_DIR = input("Choose a new output dir name: ")
#             break
#         else:
#             continue
# if not os.path.exists("output"):
#     os.mkdir("output")

output_path = os.path.join("output", OUTPUT_DIR)
# # if os.path.exists(output_path):
shutil.rmtree(output_path)
os.mkdir(output_path)


def process_spreadsheet(file_path):
    print(file_path)

    # Reads file into df then runs through the de-merger
    raw_df = pd.read_excel(file_path, "Table_0")
    df = duplicate_shift(raw_df)

    # Creates table object and checks all values
    table = Table(df=df)
    table.check_dates()
    table.check_normals(QUOTIENT_MAX=QUOTIENT_MAX)
    table.check_labels()

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
    name = os.path.basename(file_path)
    wb.save(os.path.join("Output", OUTPUT_DIR, name))



# Process all files in parallel
if __name__ == "__main__":

    start_time = time.perf_counter()
    with Pool(processes=4) as pool:
        pool.map(process_spreadsheet, file_paths)
    print(f"Spread sheet processing completed in: {time.perf_counter()-start_time}")
