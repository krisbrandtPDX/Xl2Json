# Xl2Json
## Web API to read excel files and return Json objects
* Drop spreadsheets in C:\Data => currently this is hardcoded
* api/Books => return a list of available Book objects with no sheet data
* api/Books/5 => return book 5 object populated with sheet data
* api/Sheets => return a list of Sheet objects for all books, with no row data
* api/Sheets/5 => return sheet 5 object populated with row data
* api/Rows => return list of Row objects for all sheets in all books
* api/Rows/5 => return row 5 object
* Id's are stored internally and are unique
* Spreadsheets are assumed to be tabular
