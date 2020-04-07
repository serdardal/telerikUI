using Deneme.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models.Contracts
{
    public class UnlockResponseModel
    {
        public List<TableModel> NotMergedDataCellTables { get; set; }

        public List<TableModel> MergedDataCellTables { get; set; }

        public List<TableModel> NotNullCellTables { get; set; }

        public List<TableModel> ShipParticularCellTables { get; set; }

        public List<MergeTableModel> MergedRangesTables { get; set; }
  
    }
}
