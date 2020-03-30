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

        bool AddNewCells(List<CellRecord> cellRecords);

        Task<bool> UpdateCellsAsync(List<CellRecord> addedCellRecords, List<CellRecord> changedCellRecords, List<CellRecord> deletedCellRecords);

        bool UpdateCells(List<CellRecord> addedCellRecords, List<CellRecord> changedCellRecords, List<CellRecord> deletedCellRecords);

        string GetTemplateName(string documentName);

        DateTime GetDate(string documentName);

        List<CellRecord> GetCellRecordsByDocName(string docName);

        List<string> GetSavedFileNames();
    }
}
