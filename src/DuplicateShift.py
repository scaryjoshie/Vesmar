# Imports
import pandas as pd

# Python Imports
from src.MiscUtils import keep_characters


def is_cell_label(string):
    """
    This function just finds the ratio of alphabetical characters to numbers. If that ratio exceeds a certain amount (7/3 or 70% alphabet), the cell is deemed as a label
    """
    string = str(string)
    len_chars = int(len(keep_characters(string, "abcdefghijklmnopqrstuvwxyz'()$")))
    len_nums = int(len(keep_characters(string, "1234567890")))
    if len_nums == 0:
        return True
    return (len_chars / len_nums) > (7 / 3)


# Duplicate Shift
def duplicate_shift(df):
    # Saves last column of original dataframe for later
    last_col = df.iloc[:, -1:]

    # Gets the number of rows and columns
    num_rows = df.shape[0]
    num_cols = df.shape[1]
    current_row = 0

    #
    while current_row < num_rows:
        squished_list = []
        for col in range(0, num_cols):
            cell = df.iat[current_row, col]
            if pd.isna(cell):
                cell = ""

            # Makes sure that cell is not a label
            if is_cell_label(cell):
                squished_list.append([cell])
            else:
                squished_list.append(str(cell).split())

        # Creates a "row list" which is just the current row if it was split into a dataframe (most of the time,
        # it's just a 1xn matrix, but if there is a split it will be a mxn matrix)
        row_list = []
        for index in range(0, max([len(a) for a in squished_list])):
            row = []
            for sub_list in squished_list:
                if len(sub_list) - 1 >= index:
                    row.append(sub_list[index])
                else:
                    row.append("")
            row_list.append(row)

        # Drops the current row
        df.drop([current_row], axis=0, inplace=True)

        # Adds split rows
        for row in row_list:
            df.loc[current_row - 0.5] = row
            df = df.sort_index().reset_index(drop=True)
            current_row += 1

        # Ticks counter up
        num_rows += len(row_list) - 1

    # Replaces last column with original last column, since that's generally more accurate
    df = df.iloc[:, :-1]
    df = pd.concat([df, last_col], axis=1, ignore_index=True)

    return df


if __name__ == "__main__":
    # Imports
    import glob

    # Gets paths of every xlsx file in specified directory
    FOLDER_PATH = "Tables\\1970 Mar Apr Right Tables"
    files_paths = glob.glob(f"{FOLDER_PATH}/*.xlsx")
    # Opens specific dataframe
    df = pd.read_excel(files_paths[294], "Table_0")

    dup = duplicate_shift(df=df)

    dup.to_excel("testing or archive\\294 done 2.xlsx", index=False)
