﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models.Contracts
{
    public class ExportRequestModel
    {
        public string FileName { get; set; }
        public List<ColoredCellModel> ColoredCellList { get; set; }
    }
}
