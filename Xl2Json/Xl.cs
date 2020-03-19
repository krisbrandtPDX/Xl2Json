using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xl2Json.Models;

namespace Xl2Json
{
    public class Xl
    {
        private const string XL_DIR = @"C:\Data";
        private Microsoft.Office.Interop.Excel.Application _xlApp;
        private Microsoft.Office.Interop.Excel.Workbooks _xlBooks;
        private List<Book> _books = new List<Book>();
        private List<Sheet> _sheets = new List<Sheet>();
        private List<Row> _rows = new List<Row>();

        public Xl(bool fetchData = false)
        {
            _xlApp = new Microsoft.Office.Interop.Excel.Application();
            _xlBooks = _xlApp.Workbooks;
            if (fetchData)
            {
                SetRows();
            }
        }

        #region "public methods"

        //get a list of available workbooks without sheet data
        public List<Book> GetBooks() 
        {
            SetBooks(XL_DIR);
            return _books;
        }

        //get a single book object with Sheet Data
        public Book GetBook(int bookId)
        {
            SetSheets();
            Book b = _books[bookId];
            return b;
        }

        //get list of sheet objects for all workbooks in XL_DIR => no row data
        public List<Sheet> GetSheets()
        {
            SetSheets();
            return _sheets;
        }

        //get list of sheet objects for a specific workbook
        public List<Sheet> GetSheets(int bookId)
        {
            SetSheets();
            Book b = _books[bookId];
            return b.Sheets;
        }

        //get a single sheet object with Row data
        public Sheet GetSheet(int sheetId)
        {
            SetRows();
            Sheet s = _sheets[sheetId];
            return s;
        }

        //get list of row objects for all sheets in all workbooks in XL_DIR
        public List<Row> GetRows()
        {
            SetRows();
            return _rows;
        }

        //get s single row object
        public Row GetRow(int rowId)
        {
            SetRows();
            Row r = _rows[rowId];
            return r;
        }

        #endregion

        #region "private methods"

        //populate a list of Book objects for  Excel Workbooks in a directory => books added without Sheets
        private void SetBooks(string excelDir)
        {
            if (_books.Count() == 0)
            {
                var directory = new DirectoryInfo(excelDir);
                var files = directory.GetFiles().OrderBy(f => f.Name);
                int id = _books.Count();
                foreach (var f in files)
                {
                    Book b = new Book()
                    {
                        Id = id++,
                        Path = f.FullName
                    };
                    _books.Add(b);
                }
            }
            return;
        }

        //populate list of Sheet objects for all sheets in all workbooks in XL_DIR => sheets added without Rows
        private void SetSheets()
        {
            if (_sheets.Count() == 0)
            {
                SetBooks(XL_DIR);
                foreach (Book b in _books)
                {
                    b.Sheets = GetSheets(b);
                    _sheets.AddRange(b.Sheets);
                }
            }
        }

        //populate list of Row objects for all rows in all Sheets in all workbooks in XL_DIR 
        private void SetRows()
        {
            if (_rows.Count() == 0)
            {
                SetSheets();
                foreach (Sheet s in _sheets)
                {
                    s.Rows = GetRows(s);
                    _rows.AddRange(s.Rows);
                }
            }
        }

        //get a list of Sheet objects for a workbook => sheets added without Rows
        private List<Sheet> GetSheets(Book book)
        {
            List<Sheet> sheets = new List<Sheet>();
            Microsoft.Office.Interop.Excel._Workbook xlBook = _xlBooks.Open(book.Path);
            Microsoft.Office.Interop.Excel.Sheets xlSheets = xlBook.Worksheets;
            int id = _sheets.Count();
            foreach (Microsoft.Office.Interop.Excel._Worksheet w in xlSheets)
            {
                Sheet s = new Sheet()
                {
                    Id = id++,
                    BookId = book.Id,
                    BookPath = book.Path,
                    Index = w.Index,
                    Name = w.Name
                };
                sheets.Add(s);
            }
            _xlBooks.Close();
            return sheets;
        }

        //get a list of Row objects for a sheet
        private List<Row> GetRows(Sheet sheet)
        {
            List<Row> rows = new List<Row>();
            Book book = _books.Where(b => b.Id == sheet.BookId).FirstOrDefault();
            Microsoft.Office.Interop.Excel._Workbook xlBook = _xlBooks.Open(book.Path);
            Microsoft.Office.Interop.Excel.Sheets xlSheets = xlBook.Worksheets;
            Microsoft.Office.Interop.Excel._Worksheet xlSheet = (Microsoft.Office.Interop.Excel._Worksheet)xlSheets[sheet.Index];
            int id = _rows.Count();
            int row = 2;
            while (xlSheet.get_Range(CellName(row, 1)).Value2 != null)
            {
                Row r = new Row()
                {
                    Id = id++,
                    SheetId = sheet.Id,
                    Index = row,
                    Data = GetRowData(GetRowVals(xlSheet, 1), GetRowVals(xlSheet, row++))
                };
                rows.Add(r);
            }
            _xlBooks.Close();
            return rows;
        }

        //get single row values, stored in a string list
        private List<string> GetRowVals(Microsoft.Office.Interop.Excel._Worksheet xlSheet, int row)
        {
            List<string> values = new List<string>();
            Microsoft.Office.Interop.Excel.Range cell = xlSheet.get_Range(CellName(row, 1));
            while (cell.Value2 != null)
            {
                values.Add("" + cell.Value2);
                cell = cell.Next;
            }
            return values;
        }

        //returns a dictionary object with key-value pairs for single row
        private Dictionary<string, string> GetRowData(List<string> headers, List<string> values)
        {
            Dictionary<string, string> data = new Dictionary<string, string>();
            for (int i = 0; i < headers.Count(); i++)
            {
                data.Add(headers[i], values[i]);
            }
            return data;
        }

        //return cell string name from row, col : (2,1) => "A2"
        private string CellName(int row, int col)
        {
            Char column = (char)(col + 64);
            return column + row.ToString();
        }
        #endregion 
    }
}
