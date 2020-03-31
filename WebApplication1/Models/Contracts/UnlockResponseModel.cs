using Deneme.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models.Contracts
{
    public class UnlockResponseModel
    {
        public List<UnlockedTableModel> DataCells { get; set; }

        public List<UnlockedTableModel> OnlyUnlockCells { get; set; }

        public List<UnlockedTableModel> NotNullCells { get; set; }

        public List<UnlockedTableModel> ShipParticularCells { get; set; }
        public List<UnlockedTableModel> FormulaCells { get; set; }
    }
}
