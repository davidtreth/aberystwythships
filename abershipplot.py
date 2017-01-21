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
                
        


def plotMariners(shipdict):
    pass

if __name__ == '__main__':
    """
    If invoked at the command-line
    """
    # Create the command line options parser
    parser = argparse.ArgumentParser()
    parser.add_argument("-s", "--vessels", action="store_true",
                        help="plot the earliest and latest date known for each vessel")
    parser.add_argument("-v", "--verbose", action="store_true",
                        help="verbose mode - print statements while reading data")                        
    parser.add_argument("-m", "--mariners", action="store_true",
                        help="plot the dates of every mariner")
    args = parser.parse_args()
    
    shipdict = abership.getVesselsInfo(verbose=args.verbose)
    if args.vessels:
        plotVessels(shipdict, args.verbose)
        plotVessels(shipdict, args.verbose, cy=False)
    if args.mariners:
        plotMariners(shipdict, args.verbose)
    plt.show()