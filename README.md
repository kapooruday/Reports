# Reports
Visual analysis of data trends

The python script has been custom made to generate excel sheets and bokeh reports. 
The script uses the below input files : "Incidents.xlsx" in the same folder
The script generates: "Average Age.xlsx" and "Average Age.html"
Data that is hard coded and can be changed in the script variables.

'to_date' is the end date of the trend (This is coded to today - 1)
'from_date' is the start date of the trend. This can be adjusted using the variable 'days' which is currently set to 9.

Function:
The input file should have the columns named as 'Number', 'Opened', 'Closed' which are interpreted as the Case number, Date created and Date closed respectively.
The output file contains 'Date','Avearge Age' which is the average date of all the open cases on a particular day
