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

        public async Task<bool> AddNewCellsAsync(List<CellRecord> cellRecords)
        {
            await _dataContext.CellRecords.AddRangeAsync(cellRecords);
            var added = await _dataContext.SaveChangesAsync();

            return true;
        }

        public List<CellRecord> GetCellRecordsByDocName(string docName)
        {
            var cells = _dataContext.CellRecords.Where(x => x.FileName == docName).ToList();

            return cells;
        }

        public DateTime GetDate(string documentName)
        {
            CellRecord record = _dataContext.CellRecords.FirstOrDefault(x => x.FileName == documentName);
            return record.Date;
        }

        public List<string> GetSavedFileNames()
        {
            List<string> fileNames = _dataContext.CellRecords.Select(c => c.FileName).Distinct().ToList();

            return fileNames;
        }

        public string GetTemplateName(string documentName)
        {
            CellRecord record = _dataContext.CellRecords.FirstOrDefault(x => x.FileName == documentName);
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
                foreach (CellRecord changedRecord in changedCellRecords)
                {
                    var entity = _dataContext.CellRecords.FirstOrDefault(item =>
                        item.RowIndex == changedRecord.RowIndex
                        && item.ColumnIndex == changedRecord.ColumnIndex
                        && item.FileName == changedRecord.FileName
                        && item.TableIndex == changedRecord.TableIndex);

                    if (entity != null)
                    {
                        entity.Data = changedRecord.Data;
                        _dataContext.CellRecords.Update(entity);
                    }
                }

            }

            if (deletedCellRecords.Count > 0)
            {
                foreach (CellRecord deletedRecord in deletedCellRecords)
                {
                    var entity = _dataContext.CellRecords.FirstOrDefault(item =>
                        item.RowIndex == deletedRecord.RowIndex
                        && item.ColumnIndex == deletedRecord.ColumnIndex
                        && item.FileName == deletedRecord.FileName
                        && item.TableIndex == deletedRecord.TableIndex);

                    if (entity != null)
                    {
                        _dataContext.CellRecords.Remove(entity);
                    }
                }
            }


            _dataContext.SaveChanges();
            return true;
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
                        && item.FileName == changedRecord.FileName
                        && item.TableIndex == changedRecord.TableIndex);

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
                        && item.FileName == deletedRecord.FileName
                        && item.TableIndex == deletedRecord.TableIndex);

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
