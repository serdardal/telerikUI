using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Data;
using WebApplication1.Models;

namespace WebApplication1.Services
{
    public class ExcelService : IExcelService
    {
        private readonly DataContext _dataContext;

        public ExcelService(DataContext dataContext)
        {
            _dataContext = dataContext;
        }
        public async Task<bool> AddNewCellsAsync(List<CellRecord> cellRecords)
        {
            await _dataContext.CellRecords.AddRangeAsync(cellRecords);
            var added = await _dataContext.SaveChangesAsync();

            return true;
        }

        public Task<bool> UpdateCellsAsync(List<CellRecord> cellRecords)
        {
            throw new NotImplementedException();
        }
    }
}
