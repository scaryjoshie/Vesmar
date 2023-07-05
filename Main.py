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
from src.conn import get_sheets, sheets_service
from src.TableOperations import Table
from src.MiscUtils import location_to_cell_id
from src.DuplicateShift import duplicate_shift

USE_LOCAL = False
COLOR_DICT = {
    "date": "fffffb85",  # yellow
    "normal": "fffc9dde",  # normal anomaly --> pink (#fc9dde)
    "label": "ff71bd74",  # label --> green (#71bd74)
    "row_culprit": "ff72a5f7",  # row_culprit --> dark blue (#515be0)
    "row": "ffbbd4fc",  # row --> light blue (#96b1ff)
}
fillers = {}
for color in COLOR_DICT.keys():
    temp = PatternFill(patternType="solid", fgColor=COLOR_DICT[color])
    fillers[color] = temp

# Gets paths of every xlsx file in specified directory
FOLDER_PATH = "Tables\\Label Tables"
# FOLDER_PATH = "Tables\\1970 Mar Apr Right Tables"
files = glob.glob(f"{FOLDER_PATH}/*.xlsx")
# max values that 2 cells can have when divided by one another before the row is flagged
QUOTIENT_MAX = 10


if USE_LOCAL:
    output_dir = "output13"
    os.mkdir(f"Output\\{output_dir}")
else:
    files = get_sheets()

# Runs through every file
for file in files:
    print(file)

    # Reads file into df then runs through the de-merger
    if USE_LOCAL:
        raw_df = pd.read_excel(file, "Table_0")
        df = duplicate_shift(raw_df)
    else:
        values = (
            sheets_service.spreadsheets()
            .values()
            .get(spreadsheetId=file["id"], range="Table_0")
            .execute()
            .get("values", [])
        )
        df = pd.DataFrame(values)

    # Creates table object and checks all values
    table = Table(df=df)
    table.check_dates()
    table.check_normals(quotient_max=QUOTIENT_MAX)
    table.check_labels()

    # Loads file into openpyxl
    wb = openpyxl.load_workbook(file)
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
    if USE_LOCAL:
        name = file.split("\\")[-1]
    else:
        name = file["name"]

    wb.save(f"Output\\{output_dir}\\{name}")
