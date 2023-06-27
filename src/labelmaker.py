# Imports
import glob
import pandas as pd


# Gets paths of every xlsx file in specified directory
FOLDER_PATH = "Tables\\Label Tables"
file_paths = glob.glob(f"{FOLDER_PATH}/*.xlsx")


# Opens all df objects as a list of Tables
label_list = []

# Opens every file under "labels"
for file_path in file_paths:

    # Reads the excel file into a pandas dataframe
    labels = list(pd.read_excel(file_path, "Table_0")[19])
    
    # Strips all labels in the list (removes outer spaces)
    labels = [label.strip() for label in labels]

    # Extends the total list of labels by the labels in this file
    label_list.extend(labels)

# Turns the label_list into a set. This basically just means that duplicates are removed
label_list = list(set(label_list))

# Sorts the list into alphabetical order
label_list.sort()

with open ("Resources\\labels.txt", "w") as f:
    for label in label_list:
        f.write(label + "\n")