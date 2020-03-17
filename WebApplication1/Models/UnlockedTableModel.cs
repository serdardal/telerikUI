using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Models;

namespace Deneme.Models
{
    public class UnlockedTableModel
    {
        public int TableIndex { get; set; }

        public List<FilledCellModel> CellList { get; set; }
    }
}
