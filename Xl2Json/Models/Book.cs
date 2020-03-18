using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Xl2Json.Models
{
    public class Book
    {
        public Book()
        {
            Sheets = new List<Sheet>();
        }
        public int Id { get; set; }
        public string Path { get; set; }

        public List<Sheet> Sheets { get; set; }
    }
}
