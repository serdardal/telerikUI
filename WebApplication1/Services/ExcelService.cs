using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Data;
using WebApplication1.Models;
using WebApplication1.Models.DbModels;

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

        public List<Default> GetDefaults()
        {
            var defaults = _dataContext.Defaults.ToList();

            return defaults;
        }

        public List<string> GetSavedFileNames()
        {
            List<string> fileNames = _dataContext.CellRecords.Select(c => c.FileName).Distinct().ToList();

            return fileNames;
        }

        public string GetTemplateName(string fileName)
        {
            CellRecord record = _dataContext.CellRecords.FirstOrDefault(x => x.FileName == fileName);
            if(record != null)
            {
                return record.TemplateName;
            }
            return null;
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

        public bool AddEndMarks(List<EndMark> endMarks)
        {
            if(endMarks.Count > 0)
            {
                _dataContext.EndMarks.AddRange(endMarks);
                _dataContext.SaveChanges();
            }
           
            return true;
        }

        public List<EndMark> GetEndMarksofTemplate(string templateName)
        {
            var endmarks = _dataContext.EndMarks.Where(x => x.TemplateName == templateName).ToList();

            return endmarks;
        }

        public List<CellRecord> GetAllRecords()
        {
            var records = _dataContext.CellRecords.ToList();

            return records;
        }

        public bool ClearAllRecords()
        {
            var records = _dataContext.CellRecords.ToList();

            _dataContext.RemoveRange(records);

            var deleted = _dataContext.SaveChanges();

            return deleted > 0;
        }

        public string GetLogoByName(string fileName)
        {
            var record = _dataContext.CellRecords.Where(x => x.FileName == fileName).FirstOrDefault();

            return record.Logo;
        }
    }
}
