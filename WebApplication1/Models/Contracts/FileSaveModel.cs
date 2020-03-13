using Deneme.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models
{
    public class FileSaveModel
    {
        public string TemplateName { get; set; }

        public string DocumentName { get; set; }

        public DateTime Date { get; set; }

        public List<FilledCellModel> CellList { get; set; }
    }
}
