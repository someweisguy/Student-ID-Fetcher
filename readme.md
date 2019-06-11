# Student ID Fetcher
This is a script that will look through a directory containing a university's student folders, and return the student ID associated with each student folder.

## Why?
Every semester, there is a purge of all the inactive student folders from the each advisor's _current and active students_ folder. Currently, the way this is done involves manually scrolling through the advisor's folder and comparing it to the results of an SQL query in the student data CRM. If there was a way to collect a list of all student IDs and names from each advisor's folder, the list could be compared with a query of inactive students, showing only a list of inactive student folders that need to be removed. 

This would drastically increase productivity, as it would speed up the time needed to perform this task which must be done several times a year and is prone to human error.  

## How to use
Call the main.py folder using `python main.py` on Windows or `python3 main.py` on a Linux or MacOS machine, navigate to the advisor's _current and active students_ directory, and click open. The script will run. When it complete, use the file navigator to select a location to save the output Excel file.

## How it works
The folder hierarchy must be arranged so that the advisor's _current and active students_ folder contains folders with files pertaining to each student. Each student folder contains an Excel document which outlines their academic plan, as well as containing both the student's name (in or around cell A6) and student ID (in or around cell F6).

The script walks through each student folder contained within an advisor's _current and active students_ folder and opens the most recent Excel file, looking for the student's ID, and the student's name in their respective cells. The script will iterate through each sheet to try and find the student ID and name. If it can't find them, it will try to find the name and ID in each Excel file in the student folder, checking from most recently edited, to least recently edited.

If it cannot find a name or ID in the excel files, it will make one last attempt to regex an ID from the folder name. If all else fails, it will leave whichever field it wasn't able to find blank. 

Finally, the script will save all the data it found into an Excel file so that you can perform an SQL query in the student data CRM to determine which students are inactive and therefore easily find which student files need to be moved out of the advisor's _current and active students_ folder. 