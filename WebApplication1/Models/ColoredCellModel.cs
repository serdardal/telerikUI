using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models
{
    public class ColoredCellModel
    {
        public int SheetIndex { get; set; }
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
        public string Color { get; set; }
    }
}
