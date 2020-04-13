using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models.DbModels
{
    public class EndMark
    {
        [Key]
        public Guid Id { get; set; }
        public string TemplateName { get; set; }
        public int SheetIndex { get; set; }
        public int RowIndex { get; set; }
        public int ColumnIndex { get; set; }
    }
}
