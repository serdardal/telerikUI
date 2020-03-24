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
    }
}