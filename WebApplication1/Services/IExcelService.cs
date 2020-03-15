using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Models;

namespace WebApplication1.Services
{
    public interface IExcelService
    {
        Task<bool> AddNewCellsAsync(List<CellRecord> cellRecords);

        Task<bool> UpdateCellsAsync(List<CellRecord> cellRecords);
    }
}
