Hello there Folks!!!

*This file application uses VLOOKUP for validating the excel files

*Using the setup file u can build the application that uses cx freeze

*For that right click holding shift button on the same folder as u have these files and make sure to
install the necessary packages they are openpyxl and cxfreeze

*after successfull building u can find the python.exe with necessary files 

*Please make sure to name the files with the extension of xlsx(Excel Worksheet).
And they are located in same folder(build) as the application(python.exe),lib and dll files.

*To run application click python.exe.

*If there are two .xlsx files it'll take that files and produce Result.xlsx wait till completed shown in cmd.
If there are more than two .xlsx files it'll ask you to select which files to select, 
make sure to type the relavant number of the file you want as shown in cmd hit enter
do the same for selecting second file as well and wait till the Result completed shown in cmd .
If there is one file it'll result in error.

*Make sure to match the data column wise in both the files.
For example: 
If we are having emp id in A column in Doc1.xlsx(example name), 
the emp id in Doc2.xlsx(example name) should be in A column as well.

*Don't forget to rename the sheet1 and sheet2 as the name you want as it avoid confusion between files.
For example:
If the data is coming from the database for Doc1.xlsx(example name) the default name would be sheet1 if you are copying,
So change the sheet name to DB to avoid the data conflict between files.

*This script will show true if there is a match between data highlighted in green ,
else false highlighted in red ,
and if there is no data it will show #N/A .

*You can check the RESULT in Result.xlsx 
If the Result.xlsx is already present in that folder it'll print Result in Result1.xlsx and so on.
