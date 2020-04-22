using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Models;
using WebApplication1.Models.DbModels;

namespace WebApplication1.Services
{
    public interface IExcelService
    {
        bool AddNewCells(List<CellRecord> cellRecords);

        bool UpdateCells(List<CellRecord> addedCellRecords, List<CellRecord> changedCellRecords, List<CellRecord> deletedCellRecords);

        string GetTemplateName(string fileName);

        DateTime GetDate(string fileName);

        List<CellRecord> GetCellRecordsByDocName(string fileName);

        List<string> GetSavedFileNames();

        List<Default> GetDefaults();

        bool AddEndMarks(List<EndMark> endMarks);

        List<EndMark> GetEndMarksofTemplate(string templateName);

        List<CellRecord> GetAllRecords();

        bool ClearAllRecords();
    }
}
