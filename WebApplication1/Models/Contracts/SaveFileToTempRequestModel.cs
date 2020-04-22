using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models.Contracts
{
    public class SaveFileToTempRequestModel
    {
        public string Base64 { get; set; }
        public string FileName { get; set; }
    }
}
