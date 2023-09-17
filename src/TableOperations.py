# Imports
import pandas as pd
from operator import attrgetter
import statistics
import os

# Python Imports
from src.CellOperations import Cell
from difflib import get_close_matches

"""
Input: dataframe 
1) Checks that all dates are numeric and in order (seperate code for the last date in column 20)
2) Checks that all non-label cells contain only numbers, decimals, and percent signs
3) Strips all cells of non-numeric/decimal/percent characters (doesnt count null cells), then makes sure that cells arent more than X multiple of each other (so no one cell is more than 10x another in a single row)
4) Could also put all label rows into a list, then look for small dev9iations in similarity and flag them
5) Highlights null label rows and also combined label rows
Output: dataframe cell numbers
"""


DECIMAL_THRESHOLD = 2


#
class Table:
    def __init__(self, df):
        # Sets df
        self.raw_df = df
        # To use later
        self.bad_cells = {
            "date": [],
            "row": [],
            "row_culprit": [],
            "row_small": [],
            "no_decimal": [],
            "normal": [],
            "label": [],
        }
        # Creates stuff
        self.open_resources()  # creates self.dates_format which is used for checking dates
        self.create_new_table()  # this function also creates self.df

    def open_resources(self):
        # Reads dates_formats to get a list for later use
        self.dates_format = list(
            pd.read_csv(os.path.join("Resources", "dates_format.tsv"), sep="\t")
        )

        # Reads allowed labels
        with open(os.path.join("Resources", "labels_258-299 (m).txt"), "r") as f:
            self.allowed_labels = []
            for line in f:
                self.allowed_labels.append(line.strip())
        # print(self.allowed_labels)

    def create_new_table(self):
        self.df = self.raw_df.copy(deep=True)
        self.df_list = []

        # Sets values for df
        for row_num in range(0, self.raw_df.shape[0]):
            for col_num in range(0, self.raw_df.shape[1]):
                # Decides whether the type is "date", "label", or "normal"
                if row_num == 0:
                    type = "date"
                elif col_num == 8:
                    type = "label"
                else:
                    type = "normal"

                # Creates cell object
                cell = Cell(
                    value=self.raw_df.iat[row_num, col_num], type=type, location=[row_num, col_num]
                )

                # Adds cell object to dataframe and to a list
                self.df.iloc[row_num, col_num] = cell
                self.df_list.append(cell)

    def check_dates(self):
        dates_list = list(filter(lambda x: x.type == "date", self.df_list))
        # print([date.value for date in dates_list])
        # print(self.dates_format)
        for i in range(0, min(len(dates_list), len(self.dates_format))):
            # print("LIST", len(dates_list))
            # print("FORMAT", len(self.dates_format))
            if not str(self.dates_format[i]).strip() == str(dates_list[i].value).strip():
                self.bad_cells["date"].append(dates_list[i])

    def check_normals(self, QUOTIENT_MAX):
        # Finds cells with anomalies
        normals_list = list(filter(lambda x: x.type == "normal", self.df_list))
        anomalies_list = list(filter(lambda x: x.has_anomalies == True, normals_list))
        self.bad_cells["normal"].extend(anomalies_list)

        # All these 2 lines do is get a list of the row numbers with usable values
        filtered_normals_list = list(filter(lambda x: x.has_usable_value == True, normals_list))
        row_nums = list(set([cell.location[0] for cell in filtered_normals_list]))

        # Finds rows with a more than 10x difference
        for row in row_nums:
            # Gets a list of everything in the row
            row_list = list(filter(lambda x: x.location[0] == row, filtered_normals_list))

            # Checks if a row contains decimal values row by turning the list into a string, then measuring the occurences of "." in the string
            row_as_string = "".join([str(i.value) for i in row_list])
            num_decimals = row_as_string.count(".")

            row_is_decimal_row = num_decimals >= DECIMAL_THRESHOLD

            # Adds all 0's to the bad cells list (under "row_culprit")
            zero_list = list(filter(lambda x: x.usable_value == 0, row_list))
            self.bad_cells["row_culprit"].extend(zero_list)

            # Removes all 0's from the list of all cells in the row
            row_list = list(filter(lambda x: x.usable_value != 0, row_list))
            row_is_corrupt = False

            if len(row_list) > 0:
                row_median = statistics.median([x.usable_value for x in row_list])
            for cell in row_list:
                # list_of_quotients = [cell.usable_value/x.usable_value for x in row_list]
                quotient = cell.usable_value / row_median
                # if max(list_of_quotients) > QUOTIENT_MAX:

                # If the row is a decimal row but there is no decimal in the cell:
                clause1 = row_is_decimal_row and not "." in str(cell.value)
                clause2 = cell.has_usable_value and not cell.usable_value >= 100
                if clause1 and clause2:
                    self.bad_cells["no_decimal"].append(cell)

                elif quotient > QUOTIENT_MAX:
                    """
                    print("QUOTIENT:", max(list_of_quotients))
                    print("cell value: ", cell.value)
                    print([cell.value for cell in row_list])
                    """
                    self.bad_cells["row_culprit"].append(cell)
                    # Reminds the program to add the entire row to the list of bad cells
                    row_is_corrupt = True

            if row_is_corrupt:
                self.bad_cells["row"].extend(row_list)

                # Finds and adds minimum value to list
                min_cell = min(row_list, key=lambda x: x.usable_value)
                self.bad_cells["row_small"].append(min_cell)

    def check_labels(self):
        return
        # Gets a list of all cells tagged as "label" cells
        labels_list = list(filter(lambda x: x.type == "label", self.df_list))
        # Finds all labels that are null and tags them
        null_labels = list(filter(lambda x: x.is_null == True, labels_list))
        self.bad_cells["label"].extend(null_labels)
        # Removes all null cells from labels_list
        labels_list = list(filter(lambda x: x.is_null == False, labels_list))
        # Flags labels
        flagged_labels = list(
            filter(lambda x: x.value.strip() not in self.allowed_labels, labels_list)
        )

        # Sees if there exists a close option to correct to for each flagged label
        for i in flagged_labels:
            close_matches = get_close_matches(i.value, self.allowed_labels, 1)
            if len(close_matches) > 0:
                i.correction = get_close_matches(i.value, self.allowed_labels, 1)[0]
                i.value = "{} --> {}".format(
                    i.value, get_close_matches(i.value, self.allowed_labels, 1)[0]
                )

        self.bad_cells["label"].extend(flagged_labels)
        # print([label.value for label in flagged_labels])


###################################

if __name__ == "__main__":
    # Imports
    import glob

    # Gets paths of every xlsx file in specified directory
    FOLDER_PATH = os.path.join("Tables", "1970 Mar Apr Right Tables")
    files_paths = glob.glob(f"{FOLDER_PATH}/*.xlsx")
    # Opens specific dataframe
    df = pd.read_excel(files_paths[200], "Table_0")

    # Fun stuff
    table = Table(df=df)
    table.check_labels()
    # table.CheckDates()
    table.check_normals(QUOTIENT_MAX=10)

    # print([cell.value for cell in table.bad_cells])
