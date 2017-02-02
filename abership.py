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
from __future__ import print_function
import glob
import os
import sys, imp
import datetime
import time
import io

import argparse
imp.reload(sys)
if sys.version_info[0] < 3:
    sys.setdefaultencoding('utf-8')
from openpyxl import load_workbook
from operator import itemgetter
from collections import defaultdict

def conv2unicode(text):
    if sys.version_info[0] < 3:
        text = unicode(text)
    return text

def getVesselsInfo(verbose=False):
    if verbose:
        print("Yn ddarllen y data\nReading Data")
    shipdict = defaultdict(dict)

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
        #print("\n")
        if verbose:
            print("Directory: {d}".format(d=s[0]))
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
            #print("\n")
            if verbose:
                print("Directory: {d}".format(d=f[0]))
            os.chdir(f[0])
            shipdict[f[1]] = defaultdict(dict)
            shipdict[f[1]]["Vessel Names2"] = []            
            shipdict[f[1]]["VesselIDs"] = []
            shipdict[f[1]]["File Names2"] = []
            shipdict[f[1]]["Worksheets2"] = []
            # find the Excel files, and sort by number
            filelist = glob.glob("File*.xlsx")  
            totalNfiles += len(filelist)
            filelistns = [int(x.split("_")[1].split("-")[1]) for x in filelist]
            filelistsorted = list(zip(filelist, filelistns))
            filelistsorted.sort(key=itemgetter(1))
            for i in filelistsorted:
                # loop through the Excel files
                #print("\n")
                if verbose:
                    print("File: {fn}".format(fn=i[0]))
                shipdict[f[1]][i[1]] = defaultdict(dict)                
                wb = load_workbook(filename=i[0])
                sheets = wb.get_sheet_names()
                for sheet in sheets:
                    # loop through the sheets in each workbook
                    if verbose:                        
                        print("Sheet: {sh}".format(sh=sheet))
                    sheet_ranges = wb[sheet]
                    # cell F2 contains vessel name, F4 its number and F6 the port of registry
                    shipname = sheet_ranges['F2'].value

                    shipnum = sheet_ranges['F4'].value                
                    shipport = sheet_ranges['F6'].value
                    # assume the vessel name is the same for each series
                    # should be, I think, except maybe if a vessel changes name
                    # take the first one it finds for each series
                    try:
                        assert(len(shipdict[f[1]]['Vessel Name'])>0)
                    except:
                        shipdict[f[1]]["Vessel Name"] = shipname
                        shipdict[f[1]]["VesselID"] = int(shipnum)
                    try:
                        assert(len(shipdict[f[1]]["Port"])>0)
                    except:
                        shipdict[f[1]]["Port"] = shipport
                    # add to list of all ship names in each series
                    shipdict[f[1]]["Vessel Names2"].append(shipname)
                    if shipnum:
                        try:
                            shipnum = int(shipnum)
                        except:
                            pass

                        shipdict[f[1]]["VesselIDs"].append(shipnum)
                        
                    shipdict[f[1]][i[1]][sheet]["Crewlist"] = []
                    shipdict[f[1]][i[1]][sheet]["FileName"] = i[0]
                    shipdict[f[1]]["File Names2"].append(i[0])
                    shipdict[f[1]]["Worksheets2"].append(sheet)
                    # headings are on row 8, 2 rows with example data 10-11
                    rownum = 12
                    # could put anything here, so that while loop doesn't immediately stop
                    marinername = "Matthew Trewella"
                    

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
                            #crewdict = {"name":marinername, "byear":str(marinerbyear), "age":str(marinerage),
                            #            "bplace":str(marinerbplace), "datejoin":str(marinerdatejoin),
                            #            "portjoin":str(marinerportjoin), "capacity":str(marinercap),
                            #            "dateleft":str(marinerdateleft), "portleft":str(marinerportleft)}
                            crewdict = {"name":marinername, "byear":conv2unicode(str(marinerbyear)), "age":conv2unicode(str(marinerage)),
                                        "bplace":conv2unicode(str(marinerbplace)), "datejoin":conv2unicode(str(marinerdatejoin)),
                                        "portjoin":conv2unicode(str(marinerportjoin)), "capacity":conv2unicode(str(marinercap)),
                                        "dateleft":conv2unicode(str(marinerdateleft)), "portleft":conv2unicode(str(marinerportleft))}

                                                                                                     
                            shipdict[f[1]][i[1]][sheet]["Crewlist"].append(crewdict)
                        rownum += 1
                        
                    
            #print("Sheets: {sh}".format(sh=sheets))
            os.chdir("..")
        os.chdir("..")
    os.chdir("..")
    print("\nTotal number of Excel files = {t}".format(t=totalNfiles))
    return shipdict
#print(shipdict) 

def printHTMLIntro(hfile):
   """ print opening HTML boilerplate to file """
   hfile.write(conv2unicode("""<!DOCTYPE html>
   <html>
   <head>
   <meta charset='UTF-8'>
   <title>Aberystwyth Shipping Records - National Library of Wales - Llyfrgell Genedlaethol Cymru</title>
   <link href='abership.css' rel='stylesheet' type=text/css media='all'>
   </head>
   <body>"""))
   
def printHTMLIntro2(hfile):
    """ write longer HTML boilerplate for neocities pages """
    
    htmlint = """<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<title>Earliest and latest dates for all vessels</title>
<link href="/style.css" rel="stylesheet" type="text/css" media="all">
<link href='/abership/abership.css' rel='stylesheet' type=text/css media='all'>
</head>
<body>
<div id="header"> 
      <img class="flagl"  src="/St_Piran's_Flag.png" alt="Cornish flag" />
      <div id = "pagetitle"><h2>Taklow Kernewek</h2>
        <p class="h3trans"><em>Cornish Things</em></p>
      </div>
      <img class="flagr" src="/St_Piran's_Flag.png" alt="Cornish flag" />
    </div>
    <nav class="horiz">
      <ul>
      <li><a href="/index.html">Home<br><div class="transtext">Tre</div></a></li>
      <li><a href="/NLPkernewek.html">Cornish NLP<br><div class="transtext">NLP Kernewek</div></a></li>
      <li><a href="/yethanwerin.html">Yeth an Werin map<br><div class="transtext">Map Yeth an Werin</div></a></li>
      <li><a href="/minescornwall_test_stjust.html">Mines of Cornwall maps<br><div class="transtext">Mappys Balyow Kernow</div></a></li>
      <li><a href="/KernowQGIS.html">Maps of Cornwall in Cornish<br><div class="transtext">Mappys Kernow Kernewek</div></a></li>
      <li><a href="/othermaps.html">Other Map Projects<br><div class="transtext">Mappys Erell</div></a></li>
      <li><a href="/aboutauthor.html">About me<br><div class="transtext">A-dro dhymm</div></a></li>
      </ul>
    </nav>
    <div id="bodytext">
  <h2>Aberystwyth Shipping Records.</h2>\n"""
    htmlint = conv2unicode(htmlint)
    hfile.write(htmlint)

def printHTMLClose(hfile):
    """ print closing HTML"""
    hfile.write(conv2unicode("</body></html>"))
    
def printCrewLists(shipdict, htmlout=""):
    shipnames = [shipdict[s]["Vessel Name"] for s in shipdict]
    #shipnames.sort()
    if htmlout:
        # open file and print intro of HTML file
        hfile = open(htmlout, "w")
        printHTMLIntro(hfile)
        hfile.write(conv2unicode("<h2>Total number of vessels: {n}</h2>".format(n=len(shipnames))))
        
        
    print("Number of vessels = {n}".format(n=len(shipnames)))
    #print(shipdict)
    for s in shipdict:
        if htmlout:
            hfile.write(conv2unicode("<h2>Series {s}. Vessel name: {n}</h2>".format(s=s, n=shipdict[s]["Vessel Name"])))
        for i in shipdict[s]:
            #print(i)
            if not(type(i) is int):
                continue
            for ws in shipdict[s][i]:
                
                
                #print(shipdict[s][i])

                #print(s, i, ws)
                #print(shipdict[s][i][ws])

                crewlist = shipdict[s][i][ws]["Crewlist"]
                
                if len(crewlist) > 0:
                    print("Series {s}. File name {f}. Sheet {ws}. Ship name: {n}. Ship Registry Number: {n2} Port of registry: {p}\n".format(
                    n=shipdict[s]["Vessel Name"], n2=shipdict[s]["VesselID"], p=shipdict[s]["Port"],
s=s, f=shipdict[s][i][ws]["FileName"], ws=ws))
                    print("{n:30}\t{y:10} {a:8}\t{p:20}\t{d:11} {pj:20} {c:20}\t{d2:20} {pl}".format(n="Name",
                          y="Birth Year",
                          a="Age",
                          p="Birthplace",
                          d="Date Joined",
                          pj="Port Joined",
                          c="Capacity",
                          d2="Date Left",
                          pl="Port Left"))
                if htmlout:
                    hfile.write(conv2unicode("<h3>Series {s}. File name {f}. Sheet {ws}. Ship name: {n}</h3><h4>Ship Registry Number: {n2} Port of registry: {p}</h4>".format(
                    n=shipdict[s]["Vessel Name"], n2=shipdict[s]["VesselID"], p=shipdict[s]["Port"],
s=s, f=shipdict[s][i][ws]["FileName"], ws=ws)))
                    hfile.write(conv2unicode("""<table>
                    <tr class='titlerow'><th>Name</th><th>Birth Year</th><th>Age</th>
                    <th>Birthplace</th><th>Date Joined</th><th>Port Joined</th>
                    <th>Capacity</th><th>Date Left</th><th>Port Left</th></tr>"""))
                         
                    

                for mar in crewlist:
                    #print(mar)
                    print("{n:30}\t{y:10} {a:8}\t{p:20}\t{d:11} {pj:20} {c:20}\t{d2:20} {pl}".format(n=mar["name"],
                          y=mar["byear"], a=mar["age"], p=mar["bplace"], d=mar["datejoin"], pj=mar["portjoin"],
c=mar["capacity"], d2=mar["dateleft"], pl=mar["portleft"]))
                    if htmlout:
                        hfile.write(conv2unicode("<tr><td>{n}</td><td>{y}</td><td>{a}</td><td>{p}</td><td>{d}</td><td>{pj}</td><td>{c}</td><td>{d2}</td><td>{pl}</td></tr>".format(
                        n=mar["name"], y=mar["byear"], a=mar["age"], p=mar["bplace"],
d=mar["datejoin"], pj=mar["portjoin"], c=mar["capacity"], d2=mar["dateleft"], pl=mar["portleft"])))
                print("\n")
                print(htmlout, hfile)
                if htmlout:
                    hfile.write(conv2unicode("</table>\n"))
        print("\n")
    if htmlout:
        printHTMLClose(hfile)
        hfile.close()
        
def writeCrewListsIndivHTML(shipdict):
    """ write crew lists to HTML one file per vessel """    
    for s in shipdict:
        htmlfname = "vessel{n}.html".format(n=str(s).zfill(3))
        plotfilename = "crewlist_{n}_{v}.png".format(n=str(s).zfill(3), v=shipdict[s]["Vessel Name"].replace(" ","_").replace("&","and"))
        hfile = io.open(htmlfname, "w", encoding="utf8")
        printHTMLIntro2(hfile)
        hfile.write(conv2unicode("<p><a href='../datesallvessels.html'>Back to index of all vessels</a>.</p>\n"))
        hfile.write(conv2unicode("<h2>Series {s}. Vessel name: {n}</h2>\n".format(s=s, n=shipdict[s]["Vessel Name"])))
        hfile.write(conv2unicode("<img class='crewdateplot' src='{imgfn}' />\n".format(imgfn=plotfilename)))
        for i in shipdict[s]:
            if not(type(i) is int):
                continue
            for ws in shipdict[s][i]:
                crewlist = shipdict[s][i][ws]["Crewlist"]
                hfile.write(conv2unicode("<h3>Series {s}. File name {f}. Sheet {ws}. Ship name: {n}</h3><h4>Ship Registry Number: {n2} Port of registry: {p}</h4>".format(
                    n=shipdict[s]["Vessel Name"], n2=shipdict[s]["VesselID"], p=shipdict[s]["Port"],
s=s, f=shipdict[s][i][ws]["FileName"], ws=ws)))
                hfile.write(conv2unicode("""<table>
                    <tr class='titlerow'><th>Name</th><th>Birth Year</th><th>Age</th>
                    <th>Birthplace</th><th>Date Joined</th><th>Port Joined</th>
                    <th>Capacity</th><th>Date Left</th><th>Port Left</th></tr>"""))
                for mar in crewlist:
                    hfile.write(conv2unicode("<tr><td>{n}</td><td>{y}</td><td>{a}</td><td>{p}</td><td>{d}</td><td>{pj}</td><td>{c}</td><td>{d2}</td><td>{pl}</td></tr>".format(
                    n=mar["name"], y=mar["byear"], a=mar["age"], p=mar["bplace"],
d=mar["datejoin"], pj=mar["portjoin"], c=mar["capacity"], d2=mar["dateleft"], pl=mar["portleft"])))
                hfile.write(conv2unicode("</table>\n"))
                
        printHTMLClose(hfile)
        hfile.close()
    
def checkNames(shipdict, htmlout=""):
    shipnames = [shipdict[s]["Vessel Name"] for s in shipdict]
    #shipnames.sort()
    if htmlout:
        # open file and print intro of HTML file
        hfile = open(htmlout, "w")
        printHTMLIntro(hfile)
        hfile.write(conv2unicode("<h2>Total number of vessels: {n}</h2>".format(n=len(shipnames))))
    print("Number of vessels = {n}".format(n=len(shipnames)))
    print("\nVessel names: ")
    for s in shipnames:
        print(s, end=", ")
    print("\n")  
    # now checks whether all instances of the ship name
    # are the same as that of the first one
    shipnames2 = [shipdict[s]["Vessel Names2"] for s in shipdict]
    shipnumbers = [shipdict[s]["VesselIDs"] for s in shipdict]
    shipfilenames = [shipdict[s]["File Names2"] for s in shipdict]
    shipwsnames = [shipdict[s]["Worksheets2"] for s in shipdict]
    zshipnames = list(zip(shipnames, shipnames2, shipnumbers, shipfilenames, shipwsnames))
    for z in zshipnames:
        print("{n} : {w} worksheets".format(n=z[0], w=len(z[1])))
        if htmlout:
            hfile.write("<h3>Vessel: {n} : {w} worksheets</h3>".format(n=z[0], w=len(z[1])))
        try:
            assert len(list(set(z[2]))) == 1
        except:              
            print("more than one ship registry number per series")
            print(z[2])
            if htmlout:
                hfile.write(conv2unicode("<p><em>more than one vessel registry number per series</em><br>Numbers found in spreadsheets: {ns}</p>".format(ns=z[2])))
    
        for s2, fn, ws in list(zip(z[1], z[3], z[4])):
            try:
                assert(z[0].strip() == s2.strip())
            except:
                print("ship name conflict")
                print(z[0], s2)
                if htmlout:
                    hfile.write(conv2unicode("<p><em>vessel name conflict</em><br>filename: {f} sheet: {ws}<br><strong>First worksheet</strong>: {v1} <strong>Current worksheet</strong>: {v2}</p>".format(
                    v1=z[0], v2=s2, f=fn, ws=ws)))
    if htmlout:
        printHTMLClose(hfile)
        hfile.close()                    #time.sleep(5)

def str2Date(datestr, assumefirstday=True, verbose=False, acceptguesses=False, hfile=None):
    datestr = datestr.strip()
    datestr = datestr.replace("/","-")
    if acceptguesses:
        datestr = datestr.replace("[","")
        datestr = datestr.replace("]","")
    if len(datestr.split("-")) == 2:
        if assumefirstday:            
            d = 1
        else:
            d = 28
        try:
            y, m = [int(i) for i in datestr.split("-")]
        except:
            pass
    elif len(datestr.split("-")) == 3:
        try:
            y, m, d = [int(i) for i in datestr.split("-")]
        except:
            pass
    else:
        pass
    try:
        outputdate = datetime.date(y,m,d)
    except:
        if datestr.lower() not in ["blk", "remains", "continued", "continues", "remains on board", "continuous", "still on board", "still remaining"]:        
            # don't print error message for "blk" meaning blank
            print("date string that cannot be processed: {d}".format(d=datestr))
            if hfile:
                hfile.write(conv2unicode("<p><em>date string that cannot be processed:</em> {d}</p>".format(d=datestr)))
        #time.sleep(1)
        if assumefirstday:            
            y, m, d = 1850, 1, 1
        else:
            y, m, d = 1920, 1, 1
        outputdate= datetime.date(y, m, d)
    return outputdate

def checkBoundsDate(dt, lbound = 1850, ubound = 1920, verbose=False, hfile=None):
   if dt.year < lbound or dt.year > ubound:
       print("Date {d} falls before {l} or after {u}.".format(d=dt, l=lbound, u=ubound))
       if hfile:
           hfile.write(conv2unicode("<p><em>Date {d} falls before {l} or after {u}</em></p>".format(d=dt,
                       l=lbound, u=ubound)))
       return False
   else:
       return True
  
def findDates(shipdict, verbose=False, htmlout=""):
    if htmlout:
        # open file and print intro of HTML file
        hfile = open(htmlout, "w")
        printHTMLIntro(hfile)
    else:
        hfile = None
    for s in shipdict:
        print("\nSeries {s}. Vessel Name, ID: {v} {num}".format(s=s,
              v=shipdict[s]["Vessel Name"], num=shipdict[s]["VesselID"]))
        hfile.write(conv2unicode("<h3>Series {s}. Vessel Name, ID: {v} {num}</h3>".format(s=s,
              v=shipdict[s]["Vessel Name"], num=shipdict[s]["VesselID"])))
        earliestdate = datetime.date(1920,1,1)
        latestdate = datetime.date(1850,1,1)
        for i in shipdict[s]:
            if not(type(i) is int):
                continue
            for ws in shipdict[s][i]:
                 crewlist = shipdict[s][i][ws]["Crewlist"]
                 print("File {f}. Sheet {ws}".format(f=shipdict[s][i][ws]["FileName"], ws=ws))
                 if htmlout:
                     hfile.write(conv2unicode("<p>filename {f}. sheet {ws}</p>".format(f=shipdict[s][i][ws]["FileName"], ws=ws)))
                 for mar in crewlist:
                     if verbose:
                         print("{n}, Dates: {d} {d2}".format(n=mar["name"],
                               d=mar["datejoin"], d2=mar["dateleft"]))                         
                     dt = str2Date(mar["datejoin"], assumefirstday=False, verbose=verbose, hfile=hfile)
                     if checkBoundsDate(dt, verbose=verbose, hfile=hfile):                         
                         if dt < earliestdate:
                             earliestdate = dt                                                                      
                     dt = str2Date(mar["dateleft"], assumefirstday=True, verbose=verbose, hfile=hfile)
                     if checkBoundsDate(dt, verbose=verbose, hfile=hfile):                         
                         if dt > latestdate:
                             latestdate = dt                     
        print("Vessel Name, ID, Dates: {v}, {num}, {de}, {dl}".format(v=shipdict[s]["Vessel Name"],
              num=shipdict[s]["VesselID"], de=earliestdate, dl=latestdate))
        if htmlout:
            hfile.write(conv2unicode("<p>Vessel Name, ID, Earliest and latest dates:<br>{v}, {num}, {de}, {dl}</p>".format(v=shipdict[s]["Vessel Name"],
                        num=shipdict[s]["VesselID"], de=earliestdate, dl=latestdate)))
    if htmlout:
        printHTMLClose(hfile)
        hfile.close()

if __name__ == '__main__':
    """
    If invoked at the command-line
    """
    # Create the command line options parser
    parser = argparse.ArgumentParser()
    parser.add_argument("-p", "--crewlists", action="store_true",
                        help="print all crew lists to console")
    parser.add_argument("-c", "--checkships", action="store_true",
                        help="check whether the ship names and registry number match across each 'series'")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="verbose mode - print statements while reading data")
    parser.add_argument("-d", "--dates", action="store_true",
                        help="find the earliest start dates and latest leave dates for each vessel")                        
    parser.add_argument("-m", "--html", action="store_true",
                        help="if set, save results to HTML files")
    parser.add_argument("-i", "--htmlindiv", action="store_true",
                        help="if set, save crew lists to individual HTML files")
    args = parser.parse_args()
    
    shipdict = getVesselsInfo(verbose=args.verbose)
    if args.html:
        htmlcrewlists = "aberships_crewlists.html"
        htmlcheckships = "aberships_checkvesselnames.html"
        htmlcheckdates = "aberships_checkdates.html"
    else:
        htmlcrewlists = ""
        htmlcheckships = ""
        htmlcheckdates = ""
    if args.crewlists:
        printCrewLists(shipdict, htmlout=htmlcrewlists)
    if args.checkships:
        checkNames(shipdict, htmlout=htmlcheckships)
    if args.dates:
        findDates(shipdict, verbose=args.verbose, htmlout=htmlcheckdates)
    if args.htmlindiv:
        writeCrewListsIndivHTML(shipdict)
                    
    