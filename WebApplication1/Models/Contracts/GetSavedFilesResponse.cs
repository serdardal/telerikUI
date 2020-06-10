using Deneme.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Models.DbModels;

namespace WebApplication1.Models.Contracts
{
    public class GetSavedFilesResponse
    {
        public string Base64File { get; set; }

        public List<TableModel> NotMergedDataCellTables { get; set; }

        public List<NotNullTableModel> NotNullCellTables { get; set; }

        public List<TableModel> ShipParticularCellTables { get; set; }

        public List<MergeTableModel> MergedRangesTables { get; set; }

        public List<EndMark> EndMarks { get; set; }

        public List<CustomFormattedTableModel> CustomFormattedCellTables { get; set; }
    }
}
