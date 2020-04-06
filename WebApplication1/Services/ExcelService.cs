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

        public bool AddNewCells(List<CellRecord> cellRecords)
        {
            _dataContext.CellRecords.AddRange(cellRecords);
            var added = _dataContext.SaveChanges();

            return true;
        }

        public List<CellRecord> GetCellRecordsByDocName(string fileName)
        {
            var cells = _dataContext.CellRecords.Where(x => x.FileName == fileName).ToList();

            return cells;
        }

        public DateTime GetDate(string fileName)
        {
            CellRecord record = _dataContext.CellRecords.FirstOrDefault(x => x.FileName == fileName);
            return record.Date;
        }

        public List<string> GetSavedFileNames()
        {
            List<string> fileNames = _dataContext.CellRecords.Select(c => c.FileName).Distinct().ToList();

            return fileNames;
        }

        public string GetTemplateName(string fileName)
        {
            CellRecord record = _dataContext.CellRecords.FirstOrDefault(x => x.FileName == fileName);
            return record.TemplateName;
        }

        public bool UpdateCells(List<CellRecord> addedCellRecords, List<CellRecord> changedCellRecords, List<CellRecord> deletedCellRecords)
        {
            if (addedCellRecords.Count > 0)
            {
                _dataContext.CellRecords.AddRange(addedCellRecords);
            }

            if (changedCellRecords.Count > 0)
            {
                _dataContext.CellRecords.UpdateRange(changedCellRecords);

            }

            if (deletedCellRecords.Count > 0)
            {
                _dataContext.CellRecords.RemoveRange(deletedCellRecords);
            }


            _dataContext.SaveChanges();
            return true;
        }
    }
}
