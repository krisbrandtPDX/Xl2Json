# Xl2Json
## Web API to read excel files and return Json objects
* Drop spreadsheets in C:\Data => currently this is hardcoded
* At startup, all of the sheet Data is loaded into a singleton instance of the Xl class
* Static singleton data is accessed using the api/* route
  * api/Books => return all Books fully populated with sheet and row data
  * api/Books/5 => return Book 5 object populated with sheet and row data
  * api/Sheets => return all Sheets objects for all books populated with row data
  * api/Sheets/5 => return Sheet 5 populated with row data
  * api/Rows => return list of Row objects for all sheets in all books
  * api/Rows/5 => return Row 5 object
* A second api layer creates a new Xl instance and reads the sheets at each request
* Dynamic data can be accessed using the api1/* route
  * api1/Books => return a list of available Book objects with no sheet data
  * api1/Books/5 => return book 5 object populated with sheet data
  * api1/Sheets => return a list of Sheet objects for all books, with no row data
  * api1/Sheets/5 => return sheet 5 object populated with row data
  * api1/Rows => return list of Row objects for all sheets in all books
  * api1/Rows/5 => return row 5 object
* Id's are stored internally and are unique
* Spreadsheets are assumed to be tabular
