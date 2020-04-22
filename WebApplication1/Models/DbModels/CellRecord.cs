using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models
{
    public class CellRecord
    {
        [Key]
        public Guid Id { get; set; }

        public int RowIndex { get; set; }

        public int ColumnIndex { get; set; }

        public string Data { get; set; }

        public int TableIndex { get; set; }

        public string TemplateName { get; set; }

        public string FileName { get; set; }

        public DateTime Date { get; set; }

        public string Type { get; set; }

        public string Logo { get; set; }
    }
}
