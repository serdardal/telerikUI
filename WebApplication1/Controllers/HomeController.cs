using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Deneme.Models;
using System.IO;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using WebApplication1.Models;
using WebApplication1.Models.Contracts;
using WebApplication1.Services;
using static OfficeOpenXml.ExcelWorksheet;

namespace Deneme.Controllers
{
    public class HomeController : Controller
    {
        private IExcelService _excelService;
        public HomeController(IExcelService excelService)
        {
            _excelService = excelService;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpGet("Home/GetTemplateByName/{docName}")]
        public string GetTemplateByName(string docName)
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), "Forms", docName);
            byte[] fileByteArray = System.IO.File.ReadAllBytes(path);
            string file = Convert.ToBase64String(fileByteArray);
            return "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + file;
        }

        [HttpGet("Home/GetSavedFileByName/{docName}")]
        public string GetSavedFileByName(string docName)
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), "Saves", docName);
            byte[] fileByteArray = System.IO.File.ReadAllBytes(path);
            string file = Convert.ToBase64String(fileByteArray);
            return "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + file;
        }


        [HttpGet]
        public IActionResult GetTemplateNames()
        {
            string[] excelFiles = Directory.GetFiles(Path.Combine(Directory.GetCurrentDirectory(), "Forms"), "*.xlsx")
                                     .Select(Path.GetFileName)
                                     .ToArray();

            return Ok(excelFiles);
        }

        [HttpGet]
        public IActionResult GetSavedFileNames()
        {
            string[] excelFiles = Directory.GetFiles(Path.Combine(Directory.GetCurrentDirectory(), "Saves"), "*.xlsx")
                                     .Select(Path.GetFileName)
                                     .ToArray();

            return Ok(excelFiles);
        }

        [HttpPost]
        public IActionResult GetUnlockedCells([FromBody] CellUnlockModel model)
        {
            string selectFolder = model.IsTemplate ? "Forms" : "Saves";
            string path = Path.Combine(Directory.GetCurrentDirectory(), selectFolder, model.DocumentName);
            List<UnlockedTableModel> dataCells = new List<UnlockedTableModel>();
            List<UnlockedTableModel> onlyUnlockedCells = new List<UnlockedTableModel>();

            FileInfo fi = new FileInfo(path);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheets worksheetList = excelPackage.Workbook.Worksheets;

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

                            if (!locked) {
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
            return Ok(new UnlockResponseModel { DataCells= dataCells, OnlyUnlockCells=onlyUnlockedCells});
        }

        [HttpPost]
        public async Task<IActionResult> WriteToExcelFile([FromBody] FileSaveModel model)
        {
            List<FilledTableModel> filledTables = model.TableList;
            
            string path = Path.Combine(Directory.GetCurrentDirectory(), "Forms", model.TemplateName);
            string fileName = model.TemplateName.Replace(".xlsx", "") + "_" + model.DocumentName + "_" + model.Date.ToString("dd.MM.yyyy") + ".xlsx";
            FileInfo fi = new FileInfo(path);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheets worksheets = excelPackage.Workbook.Worksheets;

                for (int k = 0; k < worksheets.Count; k++)
                {
                    List<FilledCellModel> cells = filledTables[k].CellList;

                    for (int i = 0; i < cells.Count; i++)
                    {
                        worksheets[k].Cells[cells[i].RowIndex, cells[i].ColumnIndex].Value = cells[i].Value;
                    }
                }
                

                string savePath = Path.Combine(Directory.GetCurrentDirectory(),
                    "Saves", fileName);
                FileInfo savePathFI = new FileInfo(savePath);
                excelPackage.SaveAs(savePathFI);

            }

            List<CellRecord> cellRecords = new List<CellRecord>();


            for (int i = 0; i < filledTables.Count; i++)
            {
                FilledTableModel table = filledTables[i];
                List<FilledCellModel> cells = table.CellList;

                if(cells.Count > 0)
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
                            FileName = fileName,
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

            string path = Path.Combine(Directory.GetCurrentDirectory(), "Saves", model.DocumentName);

            FileInfo fi = new FileInfo(path);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheets worksheets = excelPackage.Workbook.Worksheets;

                //excel tablosu üzerinde ekleme-değiştirme-silme işlemlerini yapar
                for (int k = 0; k < worksheets.Count; k++)
                {
                    List<FilledCellModel> cells = new List<FilledCellModel>();
                    cells.AddRange(filledTables[k].CellList);
                    cells.AddRange(changedTables[k].CellList);
                    cells.AddRange(emptiedTables[k].CellList);

                    for (int i = 0; i < cells.Count; i++)
                    {
                        worksheets[k].Cells[cells[i].RowIndex, cells[i].ColumnIndex].Value = cells[i].Value;
                    }
                }
     
                FileInfo savePathFI = new FileInfo(path);
                excelPackage.SaveAs(savePathFI);

            }

            //database e kaydetme
            List<CellRecord> addedCellRecords = new List<CellRecord>();
            List<CellRecord> updatedCellRecords = new List<CellRecord>();
            List<CellRecord> deletedCellRecords = new List<CellRecord>();

            string templateName = _excelService.GetTemplateName(model.DocumentName);
            DateTime date = _excelService.GetDate(model.DocumentName);

            for (int i = 0; i < filledTables.Count; i++)
            {
                List<FilledCellModel> cells = filledTables[i].CellList;

                if(cells.Count > 0)
                {
                    foreach(FilledCellModel cell in cells)
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

            await _excelService.UpdateCellsAsync(addedCellRecords,updatedCellRecords,deletedCellRecords);

            return Ok();
        }



        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
