using Deneme.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models
{
    public class FilledTableModel
    {
        public int TableIndex { get; set; }

        public List<FilledCellModel> CellList { get; set; }
    }
}
