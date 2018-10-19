# 1. Find the video file at this path: C:\Users\vladimir.georgiev\Documents\BT Inputs
# 2. Rename the file according to it's time stamp from an excel sheet
# 3. Copy the file after it is renamed to this location: C:\Users\vladimir.georgiev\Dropbox\Inputs BT

def printlist(liste):
    for i in liste:
        print(i)
    print()

#returns all events of the excel_data, that are on the same date as the timestamp
def searchPossibilities(excel_data, timestamp):
    possibilities=[]
    timestamp=formatTimestamp(timestamp)
    for i in range(len(excel_data)):
        if excel_data[i][0]==timestamp[0]:
            possibilities.append(excel_data[i])
    return possibilities
     
#formats timestamp to [time, date]
def formatTimestamp(timestamp):
    #timestamp format= 2018-10-18 15-12-43
    if not isinstance(timestamp, (list,)):
        timestamp=timestamp.split(" ")
        date=timestamp[0].split("-")
        date=date[2]+"."+date[1]+"."+date[0]
        time=timestamp[1].split("-")
        time=time[0]+":"+time[1]
        return [date,time]
    else:
        return timestamp

#Use settings from "settings.py" 
from settings import *

#needed for renaming and path searching
import os

#needed for Excel-interface
from openpyxl import Workbook,load_workbook

#needed for moving a file.. Use shutil.move(from, to)
import shutil

#connect to excel file
wb = load_workbook(path_excel)
ws=wb.active

print("grabbing and formatting excel data...", end="\n\n")
#get excel data
excel_data=[]
for row in ws.rows:
    excel_data.append([row[0].value, row[1].value, row[2].value])

#drop headers
excel_data=excel_data[1:]

#get missing dates
for i in range(len(excel_data)-1):
    if excel_data[i+1][0]==None:
        excel_data[i+1][0]=excel_data[i][0]

print("got following data:")
printlist(excel_data)


#iterate over searchpath
videofiles=[]
for i in os.listdir(path_input):
    if i.endswith("."+videoformat):
        videofiles.append(i)
        print("found file '{}': searching in excel_sheet...".format(i))

#printlist(searchPossibilities(excel_data, videofiles[0]))