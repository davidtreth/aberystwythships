# -*- coding: utf-8 -*-
"""

Read the National Library of Wales' Aberystwyth Shipping Records
https://github.com/LlGC-NLW/shippingrecords/releases

using openpyxl

expect a series of directories that result from unzipping
https://github.com/LlGC-NLW/shippingrecords/releases/download/v0.1/ABERSHIP_transcription_vtls004566921.zip

expects to be launched from above the directory ABERSHIP_transcription_vtls004566921
below that each directory is called "Series N+1 - N+10"
the last directory is "Series 541 - 544"

below this each series has 1 directory each
and below this are the .xlsx Excel files
which have one or more worksheets

at present the program simply loops through and prints out
the names of the subdirectories, and files, and the details of crew:

"Name", "Birth Year", "Age", "Birthplace", "Date Joined", "Port Joined", "Capacity", "Date Left", "Port Left"

"""

import glob
import os
import sys, imp
import datetime
imp.reload(sys)
if sys.version_info[0] < 3:
    sys.setdefaultencoding('utf-8')
from openpyxl import load_workbook
from operator import itemgetter


transcriptdir = glob.glob("ABERSHIP_transcription*")

if len(transcriptdir) > 1:
    print("only one ABERSHIP_transcription directory expected")
    print(transcriptdir)
else:
    transcriptdir = transcriptdir[0]
    os.chdir(transcriptdir)
    
seriesdirs = glob.glob("Series*")
# make the directories be sorted by number not alphabetically
# i.e. 11 should come before 101 etc.
seriesstarts = [int(d.split(" ")[1]) for d in seriesdirs]
#print(seriesdirs)
#print(seriesstarts)
seriesdirssorted = list(zip(seriesdirs, seriesstarts))
seriesdirssorted.sort(key=itemgetter(1))
#print(seriesdirssorted)

# count how many Excel files we have
totalNfiles = 0

for s in seriesdirssorted:
    # loop through the top-level directies
    print("\n")
    print(s[0])
    os.chdir(s[0])
    # find the second level directories (1 per ship) and sort by number
    filedirs = glob.glob("Series*")
    # the  [:3] is needed for Series_525a_vtls004574161
    fileindices = [int(d.split("_")[1][:3]) for d in filedirs]
    filedirssorted = list(zip(filedirs, fileindices))
    filedirssorted.sort(key=itemgetter(1))
    #print(filedirssorted)
    for f in filedirssorted:
        # loop through the second level directories
        print("\n")
        print(f[0])
        os.chdir(f[0])
        # find the Excel files, and sort by number
        filelist = glob.glob("File*.xlsx")  
        totalNfiles += len(filelist)
        filelistns = [int(x.split("_")[1].split("-")[1]) for x in filelist]
        filelistsorted = list(zip(filelist, filelistns))
        filelistsorted.sort(key=itemgetter(1))
        for i in filelistsorted:
            # loop through the Excel files
            print("\n")
            print(i[0])
            wb = load_workbook(filename=i[0])
            sheets = wb.get_sheet_names()
            for sheet in sheets:
                # loop through the sheets in each workbook
                print("Sheet: {sh}".format(sh=sheet))
                sheet_ranges = wb[sheet]
                # cell F2 contains vessel name, F4 its number and F6 the port of registry
                shipname = sheet_ranges['F2'].value
                shipnum = sheet_ranges['F4'].value
                shipport = sheet_ranges['F6'].value
                print("Ship name: {n}. Official Number: {n2} Port of registry: {p}\n".format(n=shipname, n2=shipnum, p=shipport))
                # headings are on row 8, 2 rows with example data 10-11
                rownum = 12
                # could put anything here, so that while loop doesn't immediately stop
                marinername = "Matthew Trewella"
                print("{n:30}\t{y:10} {a:8}\t{p:20}\t{d:11} {pj:20} {c:20}\t{d2:20} {pl}".format(n="Name",
                                                                                                y="Birth Year",
                                                                                                a="Age",
                                                                                                p="Birthplace",
                                                                                                d="Date Joined",
                                                                                                pj="Port Joined",
                                                                                                c="Capacity",
                                                                                                d2="Date Left",
                                                                                                pl="Port Left"))
                while marinername:                    
                    marinername = sheet_ranges['A{r}'.format(r=rownum)].value
                    marinerbyear= sheet_ranges['B{r}'.format(r=rownum)].value
                    if type(marinerbyear) is datetime.datetime:
                        marinerbyear = marinerbyear.date()
                    marinerage= sheet_ranges['C{r}'.format(r=rownum)].value
                    marinerbplace= sheet_ranges['D{r}'.format(r=rownum)].value
                    marinerdatejoin = sheet_ranges['J{r}'.format(r=rownum)].value
                    if type(marinerdatejoin) is datetime.datetime:
                        marinerdatejoin = marinerdatejoin.date()
                    marinerportjoin = sheet_ranges['K{r}'.format(r=rownum)].value
                    marinercap= sheet_ranges['L{r}'.format(r=rownum)].value
                    marinerdateleft = sheet_ranges['M{r}'.format(r=rownum)].value
                    if type(marinerdateleft) is datetime.datetime:
                        marinerdateleft = marinerdateleft.date()
                    marinerportleft = sheet_ranges['N{r}'.format(r=rownum)].value                   
                    if marinername:
                        print("{n:30}\t{y:10} {a:8}\t{p:20}\t{d:11} {pj:20} {c:20}\t{d2:20} {pl}".format(n=marinername,
                                                                                                           y=str(marinerbyear),
                                                                                                           a=str(marinerage),
                                                                                                           p=str(marinerbplace),
                                                                                                           d=str(marinerdatejoin),
                                                                                                           pj=str(marinerportjoin),
                                                                                                           c=str(marinercap),
                                                                                                           d2=str(marinerdateleft),
                                                                                                           pl=str(marinerportleft)))
                    rownum += 1
                        
                    
            #print("Sheets: {sh}".format(sh=sheets))
        os.chdir("..")
    os.chdir("..")
    
print("\nTotal number of Excel files = {t}".format(t=totalNfiles))
