# excel-persistent-sort
Save and restore the Sort Order of Tables in Excel. 

Sometimes Excel will reset the Sort Order, which is particularly annoying when you are sorting by several columns. This tool lets you save and restore the sort order, and these saved states persists across closing the file.

## Features
- Save the state in workbooks persistently (using CustomXML object).
- Restore saved Sort Order States.
- Partially restore any Sort Order State to a table if at least one column is present.
- Re-associate orphaned Sort Order States (i.e., Table name changed).

## Screenshots
![Screenshot of tool in action](images/Screenshot01.PNG)