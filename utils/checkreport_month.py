# Check the latest month, and run for the latest month
import os 
import re

import datetime
from tkinter.tix import FileSelectBox
from unittest import skip

from utils.config import directory

def wantedFileName():
    """
    Return the wanted filename of the Excel Raw file, based on the max month
    and the max month,
    and the max month datetime object.
    """
    dirr = directory()

    rawfiles = os.listdir(dirr.raw_dir)
    # wantedfiles = [file for file in rawfiles if "Allocation Malaysia" in file]
    # fileMonths = [re.search(r'\((.*?)\)', file).group(1) for file in wantedfiles]

    files = []
    for file in rawfiles:
        try:
            if "Allocation Malaysia" in file :
                fileMonths = re.search(r'\((.*?)\)', file).group(1)
                files.append(fileMonths)
        except AttributeError: 
            pass


    dateLists = [datetime.datetime.strptime(date, "%b'%y") for date in files]
    max_date_index = dateLists.index(max(dateLists))

    wanted_filename = rawfiles[max_date_index]
    max_month = files[max_date_index]
    max_month_datetime = dateLists[max_date_index]
    
    return wanted_filename, max_month, max_month_datetime

if __name__ == "__main__":
    print(wantedFileName())