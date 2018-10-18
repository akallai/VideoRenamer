# 1. Find the video file at this path: C:\Users\vladimir.georgiev\Documents\BT Inputs
# 2. Rename the file according to it's time stamp from an excel sheet
# 3. Copy the file after it is renamed to this location: C:\Users\vladimir.georgiev\Dropbox\Inputs BT

#Use Settings from "settings.py" 
from settings import *

#needed for renaming and path searching
import os

#needed for Excel-interface
from openpyxl import Workbook,load_workbook

#connect to excel file
wb = load_workbook(path_excel)
print(wb.sheetnames)
