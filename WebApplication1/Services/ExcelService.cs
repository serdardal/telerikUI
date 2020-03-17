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

        public DateTime GetDate(string documentName)
        {
            CellRecord record = _dataContext.CellRecords.FirstOrDefault(x => x.FileName == documentName);
            return record.Date;
        }

        public string GetTemplateName(string documentName)
        {
            CellRecord record = _dataContext.CellRecords.FirstOrDefault(x => x.FileName == documentName);
            return record.TemplateName;
        }

        public async Task<bool> UpdateCellsAsync(List<CellRecord> addedCellRecords, List<CellRecord> changedCellRecords, List<CellRecord> deletedCellRecords)
        {
            if(addedCellRecords.Count > 0)
            {
                await _dataContext.CellRecords.AddRangeAsync(addedCellRecords);
            }

            if (changedCellRecords.Count > 0)
            {
                foreach (CellRecord changedRecord in changedCellRecords)
                {
                    var entity = _dataContext.CellRecords.FirstOrDefault(item => 
                        item.RowIndex == changedRecord.RowIndex 
                        && item.ColumnIndex == changedRecord.ColumnIndex
                        && item.FileName == changedRecord.FileName);

                    if(entity != null)
                    {
                        entity.Data = changedRecord.Data;
                        _dataContext.CellRecords.Update(entity);
                    }
                }
                
            }

            if(deletedCellRecords.Count > 0)
            {
                foreach (CellRecord deletedRecord in deletedCellRecords)
                {
                    var entity = _dataContext.CellRecords.FirstOrDefault(item =>
                        item.RowIndex == deletedRecord.RowIndex
                        && item.ColumnIndex == deletedRecord.ColumnIndex
                        && item.FileName == deletedRecord.FileName);

                    if (entity != null)
                    {
                        _dataContext.CellRecords.Remove(entity);
                    }
                }
            }


            await _dataContext.SaveChangesAsync();
            return true;
        }
    }
}
