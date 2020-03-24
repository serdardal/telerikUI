using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Deneme.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using WebApplication1.Models;
using WebApplication1.Models.Contracts;
using WebApplication1.Services;

namespace WebApplication1.Controllers
{
    public class SecondPageController : Controller
    {
        private IExcelService _excelService;

        public SecondPageController(IExcelService excelService)
        {
            _excelService = excelService;
        }
        public IActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public IActionResult GetTemplateNames()
        {
            string[] excelFiles = Directory.GetFiles(Path.Combine(Directory.GetCurrentDirectory(), "Forms"), "*.xlsx")
                                     .Select(Path.GetFileName)
                                     .ToArray();

            return Ok(excelFiles);
        }

        [HttpGet("SecondPage/GetTemplateByName/{docName}")]
        public string GetTemplateByName(string docName)
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), "Forms", docName);
            byte[] fileByteArray = System.IO.File.ReadAllBytes(path);
            string file = Convert.ToBase64String(fileByteArray);
            return "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + file;
        }

        [HttpPost]
        public IActionResult GetUnlockedCells([FromBody] CellUnlockModel model)
        {
            string templateName = "";
            if (model.IsTemplate)
            {
                templateName = model.DocumentName;
            }
            else
            {
                templateName = _excelService.GetTemplateName(model.DocumentName);
            }

            string path = Path.Combine(Directory.GetCurrentDirectory(), "Forms", templateName);
            List<UnlockedTableModel> dataCells = new List<UnlockedTableModel>();
            List<UnlockedTableModel> onlyUnlockedCells = new List<UnlockedTableModel>();

            FileInfo fi = new FileInfo(path);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheets worksheetList = excelPackage.Workbook.Worksheets;

                //kayıt dosyasıysa önce cellerin içi doldurulur.
                if (!model.IsTemplate)
                {
                    List<CellRecord> savedCells = _excelService.GetCellRecordsByDocName(model.DocumentName);

                    foreach(CellRecord cell in savedCells)
                    {
                        var sheet = worksheetList[cell.TableIndex];
                        sheet.Cells[cell.RowIndex, cell.ColumnIndex].Value = cell.Data;
                    }
                }

                //data ve unlock celleri bulur.
                for (int k = 0; k < worksheetList.Count; k++)
                {
                    var currentWorksheet = worksheetList[k];

                    dataCells.Add(new UnlockedTableModel { TableIndex = k, CellList = new List<FilledCellModel>() });
                    onlyUnlockedCells.Add(new UnlockedTableModel { TableIndex = k, CellList = new List<FilledCellModel>() });

                    List<string> mergedCellList = currentWorksheet.MergedCells.ToList();
                    //kilitli olmayan ve merge edilmemiş hücreleri bulur ve listeye ekler
                    for (int i = 1; i < 300; i++)
                    {
                        for (int j = 1; j < 300; j++)
                        {
                            var currentCell = currentWorksheet.Cells[i, j];
                            bool locked = currentCell.Style.Locked;
                            bool merged = currentCell.Merge;

                            if (!locked)
                            {
                                var value = currentCell.Value;

                                if (!merged) //its data cell
                                {
                                    dataCells[k].CellList.Add(new FilledCellModel { RowIndex = i, ColumnIndex = j, Value = value == null ? null : value.ToString() });
                                }
                                else
                                {
                                    var mergeAdress = currentWorksheet.MergedCells[i, j];
                                    string masterCellName = mergeAdress.Split(":")[0];
                                    var masterCell = currentWorksheet.Cells[masterCellName];

                                    if (masterCell.Start.Row == i && masterCell.Start.Column == j) //now we are in master cell so its data cell
                                    {
                                        dataCells[k].CellList.Add(new FilledCellModel { RowIndex = i, ColumnIndex = j, Value = value == null ? null : value.ToString() });
                                    }
                                    else //its only unlock cell
                                    {
                                        onlyUnlockedCells[k].CellList.Add(new FilledCellModel { RowIndex = i, ColumnIndex = j, Value = value == null ? null : value.ToString() });
                                    }
                                }

                            }
                        }
                    }

                }


            }
            return Ok(new UnlockResponseModel { DataCells = dataCells, OnlyUnlockCells = onlyUnlockedCells });
        }

        [HttpPost]
        public async Task<IActionResult> WriteToDatabase([FromBody] FileSaveModel model)
        {
            List<FilledTableModel> filledTables = model.TableList;
            List<CellRecord> cellRecords = new List<CellRecord>();


            for (int i = 0; i < filledTables.Count; i++)
            {
                FilledTableModel table = filledTables[i];
                List<FilledCellModel> cells = table.CellList;

                if (cells.Count > 0)
                {
                    foreach (FilledCellModel cell in cells)
                    {
                        cellRecords.Add(new CellRecord
                        {
                            RowIndex = cell.RowIndex,
                            ColumnIndex = cell.ColumnIndex,
                            Data = cell.Value,
                            TableIndex = i,
                            TemplateName = model.TemplateName,
                            FileName = model.DocumentName,
                            Date = model.Date
                        });
                    }


                }

            }

            await _excelService.AddNewCellsAsync(cellRecords);

            return Ok();
        }

        [HttpPost]
        public async Task<IActionResult> UpdateExistingFile([FromBody] FileUpdateModel model)
        {
            List<FilledTableModel> filledTables = model.FilledTableList;
            List<FilledTableModel> changedTables = model.ChangedTableList;
            List<FilledTableModel> emptiedTables = model.EmptiedTableList;

            //database e kaydetme
            List<CellRecord> addedCellRecords = new List<CellRecord>();
            List<CellRecord> updatedCellRecords = new List<CellRecord>();
            List<CellRecord> deletedCellRecords = new List<CellRecord>();

            string templateName = _excelService.GetTemplateName(model.DocumentName);
            DateTime date = _excelService.GetDate(model.DocumentName);

            for (int i = 0; i < filledTables.Count; i++)
            {
                List<FilledCellModel> cells = filledTables[i].CellList;

                if (cells.Count > 0)
                {
                    foreach (FilledCellModel cell in cells)
                    {
                        addedCellRecords.Add(new CellRecord
                        {
                            RowIndex = cell.RowIndex,
                            ColumnIndex = cell.ColumnIndex,
                            Data = cell.Value,
                            TableIndex = i,
                            TemplateName = templateName,
                            FileName = model.DocumentName,
                            Date = date
                        });
                    }
                }

            }

            for (int i = 0; i < changedTables.Count; i++)
            {
                List<FilledCellModel> cells = changedTables[i].CellList;

                if (cells.Count > 0)
                {
                    foreach (FilledCellModel cell in cells)
                    {
                        updatedCellRecords.Add(new CellRecord
                        {
                            RowIndex = cell.RowIndex,
                            ColumnIndex = cell.ColumnIndex,
                            Data = cell.Value,
                            TableIndex = i,
                            TemplateName = templateName,
                            FileName = model.DocumentName,
                            Date = date
                        });
                    }
                }

            }

            for (int i = 0; i < emptiedTables.Count; i++)
            {
                List<FilledCellModel> cells = emptiedTables[i].CellList;

                if (cells.Count > 0)
                {
                    foreach (FilledCellModel cell in cells)
                    {
                        deletedCellRecords.Add(new CellRecord
                        {
                            RowIndex = cell.RowIndex,
                            ColumnIndex = cell.ColumnIndex,
                            Data = cell.Value,
                            TableIndex = i,
                            TemplateName = templateName,
                            FileName = model.DocumentName,
                            Date = date
                        });
                    }
                }

            }

            await _excelService.UpdateCellsAsync(addedCellRecords, updatedCellRecords, deletedCellRecords);

            return Ok();
        }

        [HttpGet]
        public IActionResult GetSavedFileNamesFromDB()
        {
            List<string> fileNames = _excelService.GetSavedFileNames();
            return Ok(fileNames);
        }

        [HttpGet("SecondPage/GetSavedFileByName/{docName}")]
        public string GetSavedFileByName(string docName)
        {
            string templateName = _excelService.GetTemplateName(docName);
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Forms", templateName);

            byte[] fileByteArray = { };

            List<CellRecord> cells = _excelService.GetCellRecordsByDocName(docName);

            FileInfo fi = new FileInfo(templatePath);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorkbook excelWorkBook = excelPackage.Workbook;

                foreach (CellRecord cell in cells)
                {
                    ExcelWorksheet worksheet = excelWorkBook.Worksheets[cell.TableIndex];

                    worksheet.Cells[cell.RowIndex, cell.ColumnIndex].Value = cell.Data;
                }

                fileByteArray = excelPackage.GetAsByteArray();
            }

            string file = Convert.ToBase64String(fileByteArray);
            return "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + file;
        }
    }
}