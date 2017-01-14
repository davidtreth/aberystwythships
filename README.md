# README #

Read the National Library of Wales' Aberystwyth Shipping Records
https://github.com/LlGC-NLW/shippingrecords/releases

using openpyxl Python library

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
