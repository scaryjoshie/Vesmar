import glob
import re
import pandas as pd
import string


# Gets the paths of every file in the directory
def get_all_files_in_dir(folder_path, file_suffix=".xlsx"):
    """
    Input: folder path
    Output: path to every file with a certain suffix in that folder (defauly is xlsx)
    """
    return glob.glob(f"{folder_path}/*{file_suffix}")


def keep_characters(string, allowed_characters):
    # Create a regular expression pattern from the allowed string
    pattern = f"[^{re.escape(allowed_characters)}]"

    if pd.isna(string):
        return ""
    else:
        # Remove all characters not present in the allowed string
        new_string = re.sub(pattern, "", string)
        return new_string


def location_to_cell_id(location):
    """
    [h,l]
    h: 0 --> 2
    l: 0 --> A
    """
    h = location[0] + 2
    l = string.ascii_uppercase[location[1]]
    return f"{l}{h}"


if __name__ == "__main__":
    print(location_to_cell_id([1, 0]))
