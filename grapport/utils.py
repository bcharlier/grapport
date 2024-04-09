import os
from datetime import datetime


def calculate_age(date_str, dtype=str):
    """
    Calculate age from a date string.

    Parameters:
    date_str (str): String representing the date of birth in the format "DD/MM/YYYY".

    Returns:
    int: Calculated age.
    """
    # Parse the string into a datetime object
    dob = datetime.strptime(date_str, "%d/%m/%Y")

    # Calculate the current date
    current_date = datetime.now()

    # Calculate the age
    age = current_date.year - dob.year - ((current_date.month, current_date.day) < (dob.month, dob.day))

    return dtype(age)


def short_path(path):
    normalized_path = os.path.normpath(path)
    last_two_dirs = normalized_path.split(os.path.sep)[-3:]
    return os.path.join(*last_two_dirs)
