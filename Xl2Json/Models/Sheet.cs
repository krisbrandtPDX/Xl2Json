using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Xl2Json.Models
{
    public class Sheet
    {
        public Sheet()
        {
            Rows = new List<Row>();
        }
        public int Id { get; set; }
        public int BookId { get; set; }

        public string BookPath { get; set; }
        public int Index { get; set; } //this maps to sheet index inside excel object => Id is unique key
        public string Name { get; set; }
        public List<Row> Rows { get; set; }
    }
}
