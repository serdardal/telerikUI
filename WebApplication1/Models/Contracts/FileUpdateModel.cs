using Deneme.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models.Contracts
{
    public class FileUpdateModel
    {
        public string DocumentName { get; set; }

        public List<FilledTableModel> FilledTableList { get; set; }

        public List<FilledTableModel> ChangedTableList { get; set; }

        public List<FilledTableModel> EmptiedTableList { get; set; }
    }
}
