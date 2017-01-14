# README #

Read the National Library of Wales' Aberystwyth Shipping Records
https://www.llgc.org.uk/en/collections/activities/research/nlw-data/aberystwyth-shipping-records-dataset/
https://github.com/LlGC-NLW/shippingrecords/releases

using openpyxl Python library

expect a series of directories that result from unzipping
https://github.com/LlGC-NLW/shippingrecords/releases/download/v0.1/ABERSHIP_transcription_vtls004566921.zip

This repository does not actually contain this file, which should be downloaded from the National Library of Wales github site.

The program expects to be launched from above the directory ABERSHIP_transcription_vtls004566921
Below that each directory is called "Series N+1 - N+10"
The last directory is "Series 541 - 544"

Below this each series (1 per vessel) has 1 directory each
and below this are the .xlsx Excel files
which have one or more worksheets.

At present the program simply loops through and prints out
the names of the subdirectories, and files, and the details of crew:

The following fields are currently output:
"Name", "Birth Year", "Age", "Birthplace", "Date Joined", "Port Joined", "Capacity", "Date Left", "Port Left"

The program can be used to output all the details to a single file such as
(in the Linux command line)
$ python abership.py > aberships_all.txt

redirecting to the text file aberships_all.txt

The command line can then be used to, for example find names connected with a
place:

$ grep Truro aberships_all.txt
[not all output shown]

William Clarke                	blk        42      	[Truro]             	1867-04-14  Swansea              Engineer            	1867-05-02           Liverpool
John Davies                   	blk        37      	Aberaeron           	1873-01-21  Cardiff              Boatswain           	1873-04-18           Truro
John Waters                   	blk        52      	St.Davids           	1873-01-21  Cardiff              Able Seaman         	1873-04-18           Truro
Thomas Hughes                 	blk        [24]    	Newcastle           	1873-01-21  Cardiff              Able Seaman         	1873-04-18           Truro
John McNeill                  	blk        18      	Campbeltown         	1873-01-21  Cardiff              Ordinary Seaman     	1873-04-18           Truro
John Williams                 	blk        24      	Aberaeron           	1873-04-28  Truro                Master              	Continues            None
John Davies                   	blk        37      	Aberaeron           	1873-04-28  Truro                Boatswain           	1873-06-16           Falmouth
John Walters                  	blk        52      	St.Davids, Pembroke 	1873-04-28  Truro                Able seaman         	1873-06-16           Falmouth
David Davies                  	blk        46      	Cardigan            	1873-04-28  Truro                Able seaman         	1873-06-16           Falmouth
John McNeill                  	blk        19      	Campbeltown, Argyll 	1873-04-28  Truro                Ordinary Seaman     	1873-06-16           Falmouth
Charles Libby                 	blk        19      	Truro               	1871-03-22  London               Ordinary Seaman     	1871-05-22           Aberaeron
Edwin Rogers                  	blk        22      	Truro               	1878-07-03  Runcorn              Able Seaman         	1878-10-11           Aberystwith
William Trewin                	blk        19      	Truro               	1875-09-11  Falmouth             Ordinary Seaman     	1876-01-15           Gloster
M Thomas                      	1882       blk     	Treforest Wales     	1902-01-06  Gloucester           Boy                 	1902-01-20           Truro
Thomas Teague                 	Blk        21      	Truro               	1864-11-15  Falmouth             Ordinary Seaman     	Blk                  Blk
William Pascoe                	1871       blk     	Porthleven          	Continues   blk                  Mate                	1908-11-08           Truro
