import pandas as pd
import os

# Python Imports
#from src.MiscUtils import keep_characters


# basically, this deals with small errors that occur specifically in this set (some rows are shifted up and the "79-81 E" is midding)
def simple_row_fix(df):
    # Gets the number of rows and columns
    num_rows = df.shape[0]
    num_cols = df.shape[1]
    
    # If column 14 doesnt exist, doesnt do the thing
    if 14 not in df.columns:
        return df

    # Saves last column of original dataframe for later
    first_row = df.iloc[0]
    #print(first_row)

    

    # Moves columns down
    df[11] = df[11].shift(1)
    df[12] = df[12].shift(1)

    if not '79-81 E' in df[14].values:
        # Shifts column 14 down
        df[14] = df[14].shift(1)
        

    # Replaces first row with original first row
    df.iloc[0] = first_row

    if not '79-81 E' in df[14].values:
    # Sets top right to num
        df.loc[0, 14] = '79-81 E'

    # Replaces moved-down dates with null
    df.loc[1, 11] = ""
    df.loc[1, 12] = ""

    # Returns
    return df



if __name__ == "__main__":
    # Imports
    import glob

    # Gets paths of every xlsx file in specified directory
    FILE_PATH = "Input\\1977 Feb Apr Right\\1977 Feb Apr Right 0302.xlsx"
    # Opens specific dataframe
    df = pd.read_excel(FILE_PATH, "Table_0")

    dup = simple_row_fix(df=df)

    dup.to_excel("output\\296 done 2.xlsx", index=False)
