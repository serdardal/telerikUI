using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Deneme.Models
{
    public abstract class CellBase
    {
        public int RowIndex { get; set; }

        public int ColumnIndex { get; set; }
    }
}
