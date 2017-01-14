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

"Name", "Birth Year", "Age", "Birthplace", "Date Joined", "Port Joined", "Capacity", "Date Left", "Port Left"