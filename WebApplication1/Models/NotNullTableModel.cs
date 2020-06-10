using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models
{
    public class NotNullTableModel
    {
        public int TableIndex { get; set; }

        public List<CellModelWithValue> CellList { get; set; }
    }
}
