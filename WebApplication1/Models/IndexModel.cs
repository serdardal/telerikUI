using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models
{
    public class IndexModel
    {
        public bool OpenInNewTab { get; set; }
        public string FileName { get; set; }
        public bool ReadOnly { get; set; }
    }
}
