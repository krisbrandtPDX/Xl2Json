using System.Collections.Generic;

namespace Xl2Json.Models
{
    public class Row
    {
        public Row()
        {
            Data = new Dictionary<string, string>();
        }
        public int Id { get; set; }
        public int SheetId { get; set; }
        public int Index { get; set; } //maps to row number in excel
        public Dictionary<string, string> Data { get; set; }
    }
}
