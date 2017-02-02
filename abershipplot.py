# -*- coding: utf-8 -*-
"""
Created on Sat Jan 21 19:11:08 2017

"""

from __future__ import print_function
import glob

import sys, imp
import datetime


import argparse
imp.reload(sys)
if sys.version_info[0] < 3:
    sys.setdefaultencoding('utf-8')

import matplotlib.pyplot as plt

import abership


def plotVessels(shipdict, verbose=False, cy=True):
    with plt.xkcd():
        fig, ax = plt.subplots(figsize=(16,8))
    if cy:
        plt.title("Cofnodion Llongau Aberystwyth\nwww.llgc.org.uk/cy/casgliadau/gweithgareddau/ymchwil/data-llgc/cofnodion-llongau-aberystwyth")    
    else:
        plt.title("Aberystwyth Shipping Records\nwww.llgc.org.uk/en/collections/activities/research/nlw-data/aberystwyth-shipping-records-dataset")    
    y = [1,1]
    ax.set_xlim(datetime.date(1850,1,1), datetime.date(1920,1,1))
    if cy:
        ax.set_xlabel("Blwyddyn")
        ax.set_ylabel("Rhif Cyfres")
    else:
        ax.set_xlabel("Year")
        ax.set_ylabel("Series number")
    for s in shipdict:
        earliestdate = datetime.date(1920,1,1)
        latestdate = datetime.date(1850,1,1)
        for i in shipdict[s]:
            if not(type(i) is int):
                continue
            for ws in shipdict[s][i]:
                 crewlist = shipdict[s][i][ws]["Crewlist"]
                 print("File {f}. Sheet {ws}".format(f=shipdict[s][i][ws]["FileName"], ws=ws))
                 for mar in crewlist:
#                     if verbose:
#                         print("{n}, Dates: {d} {d2}".format(n=mar["name"],
#                               d=mar["datejoin"], d2=mar["dateleft"]))                         
                     dt = abership.str2Date(mar["datejoin"], assumefirstday=False, verbose=verbose)
                     if abership.checkBoundsDate(dt):                         
                         if dt < earliestdate:
                             earliestdate = dt                                                                      
                     dt = abership.str2Date(mar["dateleft"], assumefirstday=True, verbose=verbose)
                     if abership.checkBoundsDate(dt):                         
                         if dt > latestdate:
                             latestdate = dt                     
        print("Vessel Name, ID, Dates: {v}, {num}, {de}, {dl}".format(v=shipdict[s]["Vessel Name"],
              num=shipdict[s]["VesselID"], de=earliestdate, dl=latestdate))
        if earliestdate != datetime.date(1920,1,1) and latestdate != datetime.date(1850,1,1):
            ax.plot([earliestdate, latestdate], y, "k-", lw=0.5)
        y = [i+1 for i in y]
    if cy:
        plt.savefig("Dyddiadau_pob_llong.png")
    else:
        plt.savefig("dates_all_vessels.png")
                
        
def CSSVessels(shipdict, verbose=False):
    htmlout = "datesallvessels.html"
    hfile = open(htmlout, "w")
    hfile.write(abership.conv2unicode("""<!DOCTYPE html>
<html>
<head>
<meta charset='utf-8'>
<title>Earliest and latest dates for all vessels</title>
<style type="text/css">
.shiprow{
margin:0;
display:block;}
.shipleft{
background-color:lightgray;
height:5px;
margin-top:2px;
margin-bottom:2px;
margin-left:0;
margin-right:0;
display:inline-block;}
.shipright{
background-color:lightgray;
height:5px;
margin-top:2px;
margin-bottom:2px;
margin-left:0;
margin-right:0;
display:inline-block;}
.shiprange{
background-color:black;
height:9px;
margin-left:0;
margin-right:0;
display:inline-block;}
.shiprange:hover{
background-color:red;
}
.shipname {
margin-left:10px;
margin-right:10px;
font-size:small;
display:inline-block;
width:100px;
}
</style>
</head>
<body>"""))
    hfile.write(abership.conv2unicode("\n<div class='shipsouter'>"))
    for s in shipdict:
        earliestdate = datetime.date(1920,1,1)
        latestdate = datetime.date(1850,1,1)
        for i in shipdict[s]:
            if not(type(i) is int):
                continue
            for ws in shipdict[s][i]:
                crewlist = shipdict[s][i][ws]["Crewlist"]
                print("File {f}. Sheet {ws}".format(f=shipdict[s][i][ws]["FileName"], ws=ws))
                for mar in crewlist:
#                     if verbose:
#                         print("{n}, Dates: {d} {d2}".format(n=mar["name"],
#                               d=mar["datejoin"], d2=mar["dateleft"]))                         
                     dt = abership.str2Date(mar["datejoin"], assumefirstday=False, verbose=verbose)
                     if abership.checkBoundsDate(dt):                         
                         if dt < earliestdate:
                             earliestdate = dt                                                                      
                     dt = abership.str2Date(mar["dateleft"], assumefirstday=True, verbose=verbose)
                     if abership.checkBoundsDate(dt):                         
                         if dt > latestdate:
                             latestdate = dt
        if earliestdate != datetime.date(1920,1,1) and latestdate != datetime.date(1850,1,1):
             leftbl = 1000*(earliestdate - datetime.date(1850,1,1)).total_seconds()/(datetime.date(1920,1,1) - datetime.date(1850,1,1)).total_seconds()
             midbl = 1000*(latestdate-earliestdate).total_seconds()/(datetime.date(1920,1,1) - datetime.date(1850,1,1)).total_seconds()
             rightbl = 1000*(datetime.date(1920,1,1) - latestdate).total_seconds()/(datetime.date(1920,1,1) - datetime.date(1850,1,1)).total_seconds()
             #imglink = "crewlists/crewlist_{n}_{v}.png".format(n=str(s).zfill(3), v=shipdict[s]["Vessel Name"].replace(" ","_").replace("&","and"))
             imglink = "crewlists/vessel{n}.html".format(n=str(s).zfill(3))
             hfile.write(abership.conv2unicode("\n<div class='shiprow'><div class='shipname'><a href='{i}'>{t}</a></div><div class='shipleft' style='width: {l:.0f}px'></div><a href='{i}'><div class='shiprange' style='width: {m:.0f}px' title='{t}'></div></a><div class='shipright' style='width: {r:.0f}px'></div></div>".format(l=leftbl, m=midbl, r=rightbl, t=shipdict[s]["Vessel Name"], i=imglink)))
    hfile.write(abership.conv2unicode("</div>\n</body></html>"))
    
def plotMariners(shipdict, verbose=False, cy=True):
    with plt.xkcd():        
        for s in shipdict:                   
            # one plot per vessel
            fig, ax = plt.subplots(figsize=(12,8))
            plt.title(shipdict[s]["Vessel Name"])
            n = 1
            ax.set_yticks([])
            if cy:
                ax.set_xlabel("Dyddiad")
            else:
                ax.set_xlabel("Date")
            for i in shipdict[s]:
                # loop through each file
                # i should be the numerical index of each file
                # otherwise skip
                if not(type(i) is int):
                    continue                
                for ws in shipdict[s][i]:
                    # loop through each worksheet                    
                    crewlist = shipdict[s][i][ws]["Crewlist"]
                    print("File {f}. Sheet {ws}".format(f=shipdict[s][i][ws]["FileName"], ws=ws))
                    for mar in crewlist:
                        # loop through the mariners in the crew list
                        # pick out the dates of joining and leaving
                        dtj = abership.str2Date(mar["datejoin"], assumefirstday=False, verbose=verbose)                                                                                                                    
                        dtl = abership.str2Date(mar["dateleft"], assumefirstday=True, verbose=verbose)
                        # check dates are within bounds and are not the defaults
                        # 1 Jan 1850 / 1 Jan 1920
                        # which indicate an unparsable strring
                        if abership.checkBoundsDate(dtj) and abership.checkBoundsDate(dtl) and dtl != datetime.date(1850,1,1) and dtj != datetime.date(1920,1,1):
                            plt.plot([dtj,dtl], [n, n], "k-")
                            n += 1
            
            # set y axis limits, and save figure to a file
            ax.set_ylim([0,n+1])
            plt.savefig("crewlist_{n}_{v}.png".format(n=str(s).zfill(3), v=shipdict[s]["Vessel Name"].replace(" ","_").replace("&","and")))


                        
def plotAllMariners(shipdict, verbose=False, cy=True):
    # for the plot with all mariners from all vessels
    # use default styles
    plt.rcdefaults()
    # plotting all mariners on a single plot
    fig, ax = plt.subplots(figsize=(16,12))
    n = 1
    ax.set_yticks([])
    ax.set_xlim(datetime.date(1850,1,1), datetime.date(1920,1,1))
    if cy:
        plt.title("Pob Llong")
    else:
        plt.title("All Vessels")
    for s in shipdict:
        for i in shipdict[s]:
            if not(type(i) is int):
                continue
            for ws in shipdict[s][i]:
                crewlist = shipdict[s][i][ws]["Crewlist"]
                print("File {f}. Sheet {ws}".format(f=shipdict[s][i][ws]["FileName"], ws=ws))
                for mar in crewlist:
                    # loop through the mariners in the crew list
                    # pick out the dates of joining and leaving
                    dtj = abership.str2Date(mar["datejoin"], assumefirstday=False, verbose=verbose)                                                                                                                    
                    dtl = abership.str2Date(mar["dateleft"], assumefirstday=True, verbose=verbose)
                    # check dates are within bounds and are not the defaults
                    # 1 Jan 1850 / 1 Jan 1920
                    # which indicate an unparsable strring
                    if abership.checkBoundsDate(dtj) and abership.checkBoundsDate(dtl) and dtl != datetime.date(1850,1,1) and dtj != datetime.date(1920,1,1):
                        plt.plot([dtj,dtl], [n, n], "k-")
                        n += 1
    # for the all in one plot
    # set y axis limits, and save figure to a file
    ax.set_ylim([0,n+1])
    plt.savefig("all_vessels_crewlists.png")
    
if __name__ == '__main__':
    """
    If invoked at the command-line
    """
    # Create the command line options parser
    parser = argparse.ArgumentParser()
    parser.add_argument("-s", "--vessels", action="store_true",
                        help="plot the earliest and latest date known for each vessel")
    
    parser.add_argument("-d", "--htmlvessels", action="store_true",
                        help="plot the earliest and latest date known for each vessel using HTML+CSS <div>s")                        
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="verbose mode - print statements while reading data")                        
    parser.add_argument("-m", "--mariners", action="store_true",
                        help="plot the dates of every mariner")
    args = parser.parse_args()
    
    shipdict = abership.getVesselsInfo(verbose=args.verbose)
    if args.vessels:
        plotVessels(shipdict, args.verbose)
        plotVessels(shipdict, args.verbose, cy=False)
    if args.htmlvessels:
        CSSVessels(shipdict, args.verbose)        
    if args.mariners:
        plotMariners(shipdict, args.verbose)
        plotAllMariners(shipdict, args.verbose)
    #plt.show()