# Imports
import pandas as pd

# Python Imports
from src.MiscUtils import keep_characters


#
class Cell:
    def __init__(self, value, type, location):
        # Value
        self.raw_value = value
        self.value = str(value)
        self.location = location

        self.correction = None

        # Types
        allowed_types = ["date", "label", "normal"]
        if type not in allowed_types:
            raise ValueError("Invalid sim type. Expected one of: %s" % allowed_types)
        else:
            self.type = type

        # Creates lists to be accessed later
        self.has_anomalies = self.check_for_anomalies()
        self.has_usable_value = self.check_for_usable_value()
        self.usable_value = self.get_usable_value()
        self.is_null = pd.isna(self.raw_value)

    def check_for_anomalies(self):
        # If the cell is null, there are no anomalies and so it returns false
        if pd.isna(self.raw_value):
            return False
        # If filtering (only keeping the specified string) changes the value, returns True
        value = str(self.value).replace("Nil", "")  # removes "Nil" from string
        filtered_value = keep_characters(value, "1234567890.%-")
        return str(filtered_value) != str(value).strip()

    def check_for_usable_value(self) -> bool:
        # If the cell is null, there are no usable values
        if pd.isna(self.raw_value):
            return False
        # Gets the cell value but removes any characters that aren't decimals or numbers
        filtered_value = keep_characters(self.value, "1234567890.")
        # Tries to turn the cell into a float. If it can't, the cell is not usable
        try:
            float(filtered_value)
            return True
        except ValueError:
            return False

    def get_usable_value(self):
        if not self.check_for_usable_value():
            return None
        else:
            filtered_value = keep_characters(self.value, "1234567890.")
            return float(filtered_value)
