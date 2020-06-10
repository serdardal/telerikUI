using Deneme.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Models.DbModels;

namespace WebApplication1.Models
{
    public class CellGroupModel
    {
        public List<TableModel> DataCellTables { get; set; }

        public List<TableModel> NotMergedDataCellTables { get; set; }

        public List<NotNullTableModel> NotNullCellTables { get; set; }

        public List<TableModel> ShipParticularCellTables { get; set; }

        public List<MergeTableModel> MergedRangesTables { get; set; }

        public List<EndMark> EndMarks { get; set; }

        public List<CustomFormattedTableModel> CustomFormattedCellTables { get; set; }

        public List<TableModel> FormulaCellTables { get; set; }
    }
}
