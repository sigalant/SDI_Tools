# SDI_Tools

## Budget Formatting Tool

### How to Use

To use the Budget formatting tool, first launch the application by double clicking the icon.

![Example of File Explorer...](./html images/icon)

A window with three button will appear. First click the button labeled "select input file", and a file explorer window will appear. 
From this window, select the Excel file you wish to format. After you will see the filepath displayed in the application window.

![Example of Application Icon](./html images/Application_Window.PNG)

Next, click on the button labeled "select output folder", and a file explorer window will appear.
From this window, navigate to the location you want the formatted Excel file to be located, and press the "select folder" button in the bottom right of the window.

![Example of Application Window with Input File Path](./html images/FileExplorer_Window)

Once the previous two steps have been completed, you may press the button labeled "format file"(If either of the two previous steps were skipped, text will appear at the bottom of the window explaining that a field is missing.). 
A file will then appear at the previously selected location with "_formatted" appended to the name of the original file. 
If a header is omitted or misspelled from the original exported AQ Excel file, a message will appear at the bottom of the application window with a warning, and the corresponding columns in the formated file will be empty. 

*[Warning]* Possible errors could occur if:
	- the first row of the exported AQ Excel file is not populated with the headers of the columns. 
	- the logo located in the "AutoQuotes Budget Script" folder is removed, or replaced with an image with a different name
