# 1. Find the video file at this path: C:\Users\vladimir.georgiev\Documents\BT Inputs
# 2. Rename the file according to it's time stamp from an excel sheet
# 3. Copy the file after it is renamed to this location: C:\Users\vladimir.georgiev\Dropbox\Inputs BT

def printlist(liste):
    for i in liste:
        print(i)
    print()

def log(logfile, logmessage):
    print(logmessage,end="")
    logfile.write(logmessage)

#1 if within time window, 0 if not
def withinTimewindow(timewindow, time, officialTime):
    time=time[:5]
    officialTime=officialTime[:5]
    tmp=time.split(":")
    inminute_time=int(tmp[0])*60+int(tmp[1])
    tmp=officialTime.split(":")
    inminute_officialTime=int(tmp[0])*60+int(tmp[1])
    #verspätet
    if inminute_time>inminute_officialTime:
        diff=inminute_time-inminute_officialTime
        if diff < timewindow[1]:
            return 1
        else:
            return 0
    #zu früh
    else:
        diff=inminute_officialTime-inminute_time
        if diff < timewindow[0]:
            return 1
        else:
            return 0  
    
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
        if timestamp.endswith("."+videoformat):
            timestamp=timestamp.replace("."+videoformat,"")
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

#needed for checking the video duration
from mutagen.mp4 import MP4

#need for sheduling 
import schedule
import time
import datetime

def job():
#start logging
    logfile=open("logfile.txt", "a")
    log(logfile, "\n__________starting new run at {}__________\n".format(datetime.datetime.now()))
    
    #connect to excel file
    wb = load_workbook(path_excel)
    ws=wb.active

    log(logfile,"grabbing and formatting excel data...\n")
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
    for i in os.listdir(path_input):
        if i.endswith("."+videoformat):
            print("\nfound file '{}': searching in excel_sheet...".format(i))
            possibilities=searchPossibilities(excel_data, i)
            #if only one result after filtering, move
            if len(possibilities)>=1:
                compensate_p=0
                for p in range(len(possibilities)):
                    if withinTimewindow(timewindow,formatTimestamp(i)[1],possibilities[p-compensate_p][1])==0:
                        del possibilities[p-compensate_p]
                        compensate_p+=1
                if len(possibilities)==1:
                    if MP4(i).info.length/60>trashconfiguration:
                        #print(os.path.join(path_output,str(possibilities[p-compensate_p][2])+"."+videoformat), "!!!!!!!")
                        if os.path.isfile(os.path.join(path_output,str(possibilities[p-compensate_p][2])+"."+videoformat)):
                            log(logfile, "{} exists already... not renaming and moving it\n".format(os.path.join(path_output,str(possibilities[p-compensate_p][2])+"."+videoformat)))
                        else:
                            shutil.move(i, os.path.join(path_output, str(possibilities[p-compensate_p][2])+"."+videoformat))
                            log(logfile,"rename {} --> {}\n".format(i,str(possibilities[p-compensate_p][2])+"."+videoformat))
                    else:
                        log(logfile, "The duration of {} is below {} minutes... its getting ignored\n".format(i, trashconfiguration))
                else:
                    log(logfile, "To many results found for {}\n".format(i))
            elif len(possibilities)==0:
                log(logfile, "did not find excel entry for {}\n".format(i))
    print("ending job at {}".format(datetime.datetime.now()))
    print("doing job again in {} minutes".format(automaticallyRun))
    logfile.close()

job()
schedule.every(automaticallyRun).minutes.do(job)
while True:
    schedule.run_pending()
    time.sleep(1)