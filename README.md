# excel-persistent-sort
Save and restore the Sort Order of Tables in Excel. 

Sometimes Excel will reset the Sort Order, which is particularly annoying when you are sorting by several columns. This tool lets you save and restore the sort order, and these saved states persists across closing the file.

## Features
- Save the state in workbooks persistently (using CustomXML object).
- Restore saved Sort Order States.
- Partially restore any Sort Order State to a table if at least one column is present.
- Re-associate orphaned Sort Order States (i.e., Table name changed).

## Notes
- The state of the Sort Order for a ListObject is stored by recording a semicolon separated list of the column names (encoded in Base64 to avoid having to escape characters), which is separated with the Sort Order)
- Only `SortOnValues` is currently supported, and only for ascending and descending (no manual sort lists).

## Screenshots
![Screenshot of tool in action](images/Screenshot01.PNG)

# Reference
- [Sort.SortFields property (Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.sort.sortfields)
- [XlSortOrder enumeration (Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.xlsortorder)
- [XlSortOn enumeration (Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.xlsorton)