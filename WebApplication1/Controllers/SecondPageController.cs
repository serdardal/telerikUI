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
            Dictionary<string, string> variableDictionary = new Dictionary<string, string>();
            variableDictionary.Add("{VesselName}", "Jean Pierre A");
            variableDictionary.Add("{BuilderNo}", "JP-01");
            variableDictionary.Add("{SerialNo}", "SN-JP-000999");
            variableDictionary.Add("{IMONo}", "9379351");
            variableDictionary.Add("{Company}", "Arkas Holding");

            UnlockResponseModel unlockResponseModel = FindUnlockedCells(new CellUnlockModel { DocumentName = docName, IsTemplate = true });
            List<UnlockedTableModel> shipParticularCells = unlockResponseModel.ShipParticularCells;

            byte[] fileByteArray = { };
            using (ExcelPackage excelPackage = GetExcelPackageByTeplateName(docName))
            {
                ExcelWorksheets worksheetList = excelPackage.Workbook.Worksheets;

                for (int i = 0; i < shipParticularCells.Count; i++)
                {
                    ExcelWorksheet templateWorksheet = worksheetList[i];
                    UnlockedTableModel table = shipParticularCells[i];
                    List<FilledCellModel> cellList = table.CellList;

                    foreach (FilledCellModel cell in cellList)
                    {
                        string key = templateWorksheet.Cells[cell.RowIndex, cell.ColumnIndex].Value.ToString();
                        if (variableDictionary.ContainsKey(key))
                        {
                            templateWorksheet.Cells[cell.RowIndex, cell.ColumnIndex].Value = variableDictionary[key];
                        }
                    }
                }

                fileByteArray = excelPackage.GetAsByteArray();
            }

            string file = Convert.ToBase64String(fileByteArray);
            return "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + file;
        }

        private UnlockResponseModel FindUnlockedCells(CellUnlockModel model)
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
            List<UnlockedTableModel> notNullCells = new List<UnlockedTableModel>();
            List<UnlockedTableModel> shipParticularCells = new List<UnlockedTableModel>();

            FileInfo fi = new FileInfo(path);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheets worksheetList = excelPackage.Workbook.Worksheets;

               
                //data ve unlock celleri bulur.
                for (int k = 0; k < worksheetList.Count; k++)
                {
                    var currentWorksheet = worksheetList[k];

                    dataCells.Add(new UnlockedTableModel { TableIndex = k, CellList = new List<FilledCellModel>() });
                    onlyUnlockedCells.Add(new UnlockedTableModel { TableIndex = k, CellList = new List<FilledCellModel>() });
                    notNullCells.Add(new UnlockedTableModel { TableIndex = k, CellList = new List<FilledCellModel>() });
                    shipParticularCells.Add(new UnlockedTableModel { TableIndex = k, CellList = new List<FilledCellModel>() });

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
                                string format = currentCell.Style.Numberformat.Format;

                                //not null cellerin belirlenmesi
                                if (value != null && value.ToString() == "{NN}")
                                {
                                    notNullCells[k].CellList.Add(new FilledCellModel { RowIndex = i, ColumnIndex = j, Value = value == null ? null : value.ToString(), Format = format });
                                }
                                //ship particular cellerin belirlenmesi
                                if (value != null && value.ToString() != "{NN}" && value.ToString().StartsWith("{") && value.ToString().EndsWith("}"))
                                {
                                    shipParticularCells[k].CellList.Add(new FilledCellModel { RowIndex = i, ColumnIndex = j, Value = value == null ? null : value.ToString(), Format = format });
                                }
                                if (format == "General")
                                {
                                    format = null;
                                }

                                if (!merged) //its data cell
                                {
                                    dataCells[k].CellList.Add(new FilledCellModel { RowIndex = i, ColumnIndex = j, Value = value == null ? null : value.ToString(), Format = format });
                                }
                                else
                                {
                                    var mergeAdress = currentWorksheet.MergedCells[i, j];
                                    string masterCellName = mergeAdress.Split(":")[0];
                                    var masterCell = currentWorksheet.Cells[masterCellName];

                                    if (masterCell.Start.Row == i && masterCell.Start.Column == j) //now we are in master cell so its data cell
                                    {
                                        dataCells[k].CellList.Add(new FilledCellModel { RowIndex = i, ColumnIndex = j, Value = value == null ? null : value.ToString(), Format = format });
                                    }
                                    else //its only unlock cell
                                    {
                                        onlyUnlockedCells[k].CellList.Add(new FilledCellModel { RowIndex = i, ColumnIndex = j, Value = value == null ? null : value.ToString(), Format = format });
                                    }
                                }

                            }
                        }
                    }

                    //kayıt dosyasıysa önce cellerin içi doldurulur.
                    if (!model.IsTemplate)
                    {
                        List<CellRecord> savedCells = _excelService.GetCellRecordsByDocName(model.DocumentName);

                        foreach (CellRecord cell in savedCells)
                        {
                            var sheet = worksheetList[cell.TableIndex];
                            sheet.Cells[cell.RowIndex, cell.ColumnIndex].Value = cell.Data;
                        }
                    }

                }


            }

            return new UnlockResponseModel { DataCells = dataCells, OnlyUnlockCells = onlyUnlockedCells, NotNullCells = notNullCells, ShipParticularCells=shipParticularCells };
        }

        [HttpPost]
        public IActionResult GetUnlockedCells([FromBody] CellUnlockModel model)
        {
            UnlockResponseModel response = FindUnlockedCells(model);
            return Ok(response);
        }

        public ActionResult SaveFileToTemp(string contentType, string base64, string fileName)
        {
            var fileContents = Convert.FromBase64String(base64);
            System.IO.File.WriteAllBytes(Directory.GetCurrentDirectory() + $"\\Temp\\{fileName}.xlsx", fileContents);

            SyncDataWithDB(fileName);

            return View("Index");
        }

        private void SyncDataWithDB(string docName)
        {
            var cells = _excelService.GetCellRecordsByDocName(docName);
            if(cells.Count > 0)// kayıt bulunuyor yani update işlemi
            {
                UpdateExistingFileInDB(docName);
            }
            else // kayıt yok yani ekleme işlemi
            {
                AddNewRecordsToDB(docName);
            }
        }

        private void UpdateExistingFileInDB(string docName)
        {
            string templateName = FindTemplateNameFromFileName(docName);
            DateTime date = FindDateFromFileName(docName);
            List<UnlockedTableModel> dataTables = FindUnlockedCells(new CellUnlockModel { DocumentName = docName, IsTemplate = false }).DataCells;

            List<CellRecord> newCellRecords = new List<CellRecord>();
            List<CellRecord> updatedCellRecords = new List<CellRecord>();
            List<CellRecord> deletedCellRecords = new List<CellRecord>();

            string tempFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Temp", docName + ".xlsx");
            FileInfo fi = new FileInfo(tempFilePath);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheets worksheetList = excelPackage.Workbook.Worksheets;

                foreach (UnlockedTableModel table in dataTables)
                {
                    ExcelWorksheet tempWorksheet = worksheetList[table.TableIndex];

                    List<FilledCellModel> cellList = table.CellList;

                    foreach (FilledCellModel cell in cellList)
                    {
                        var tempCell = tempWorksheet.Cells[cell.RowIndex, cell.ColumnIndex].Value;

                        if (tempCell != null) // eklenmiş veya update edilmiş olabilir
                        {
                            string value = tempCell.ToString();

                            if(cell.Value == null) //kayit yoksa yeni eklemiştir
                            {
                                newCellRecords.Add(new CellRecord
                                {
                                    RowIndex = cell.RowIndex,
                                    ColumnIndex = cell.ColumnIndex,
                                    Data = value,
                                    TableIndex = table.TableIndex,
                                    TemplateName = templateName,
                                    FileName = docName,
                                    Date = date
                                });

                            }else if (cell.Value != value) // kayıt var ve temptekinden farklıysa update edilmiştir.
                            {
                                updatedCellRecords.Add(new CellRecord
                                {
                                    RowIndex = cell.RowIndex,
                                    ColumnIndex = cell.ColumnIndex,
                                    Data = value,
                                    TableIndex = table.TableIndex,
                                    TemplateName = templateName,
                                    FileName = docName,
                                    Date = date
                                });
                            }

                            
                        }
                        else // silinmiş olabilir.
                        {
                            if(cell.Value != null) // önceden kayıt varsa silinmiştir
                            {
                                deletedCellRecords.Add(new CellRecord
                                {
                                    RowIndex = cell.RowIndex,
                                    ColumnIndex = cell.ColumnIndex,
                                    Data = null,
                                    TableIndex = table.TableIndex,
                                    TemplateName = templateName,
                                    FileName = docName,
                                    Date = date
                                });
                            }
                        }
                    }
                }
            }

            _excelService.UpdateCells(newCellRecords, updatedCellRecords, deletedCellRecords);
        }

        private void AddNewRecordsToDB(string docName)
        {
            string templateName = FindTemplateNameFromFileName(docName);
            DateTime date = FindDateFromFileName(docName);
            List<UnlockedTableModel> dataTables = FindUnlockedCells(new CellUnlockModel { DocumentName = templateName, IsTemplate = true }).DataCells;

            List<CellRecord> newCellRecords = new List<CellRecord>();

            string tempFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Temp", docName + ".xlsx");
            FileInfo fi = new FileInfo(tempFilePath);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheets worksheetList = excelPackage.Workbook.Worksheets;

                foreach(UnlockedTableModel table in dataTables)
                {
                    ExcelWorksheet tempWorksheet = worksheetList[table.TableIndex];

                    List<FilledCellModel> cellList = table.CellList;

                    foreach(FilledCellModel cell in cellList)
                    {
                        var tempCell = tempWorksheet.Cells[cell.RowIndex, cell.ColumnIndex].Value;
                        
                        if (tempCell != null)
                        {
                            string value = tempCell.ToString();
                            newCellRecords.Add(new CellRecord
                            {
                                RowIndex = cell.RowIndex,
                                ColumnIndex = cell.ColumnIndex,
                                Data = value,
                                TableIndex = table.TableIndex,
                                TemplateName = templateName,
                                FileName = docName,
                                Date = date
                            });
                        }
                    }
                }

            }

            _excelService.AddNewCells(newCellRecords);
        }

        private string FindTemplateNameFromFileName(string fileName)
        {
            List<string> templateNames = Directory.GetFiles(Path.Combine(Directory.GetCurrentDirectory(), "Forms"), "*.xlsx")
                .Select(Path.GetFileName).ToList();

            foreach (string template in templateNames)
            {
                string name = template.Replace(".xlsx", "");

                if (fileName.StartsWith(name))
                {
                    return template;
                }
            }

            return String.Empty;
        }

        private DateTime FindDateFromFileName(string fileName)
        {
            return DateTime.Parse(fileName.Substring(fileName.Length - 10, 10));
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
            byte[] fileByteArray = { };

            using (ExcelPackage excelPackage = GetSavedExcelPackageByName(docName))
            {
                fileByteArray = excelPackage.GetAsByteArray();
            }

            string file = Convert.ToBase64String(fileByteArray);
            return "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + file;
        }

        private ExcelPackage GetSavedExcelPackageByName(string docName)
        {
            string templateName = _excelService.GetTemplateName(docName);
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Forms", templateName);

            List<CellRecord> cells = _excelService.GetCellRecordsByDocName(docName);

            FileInfo fi = new FileInfo(templatePath);
            ExcelPackage excelPackage = new ExcelPackage(fi);

            ExcelWorkbook excelWorkBook = excelPackage.Workbook;

            foreach (CellRecord cell in cells)
            {
                ExcelWorksheet worksheet = excelWorkBook.Worksheets[cell.TableIndex];

                worksheet.Cells[cell.RowIndex, cell.ColumnIndex].Value = cell.Data;
            }

            return excelPackage;
        }

        private ExcelPackage GetExcelPackageByTeplateName(string templateName)
        {
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Forms", templateName);
            FileInfo fi = new FileInfo(templatePath);
            ExcelPackage excelPackage = new ExcelPackage(fi);

            return excelPackage;
        }

        [HttpGet("SecondPage/GetProtectedSavedFileByName/{docName}")]
        public string GetProtectedSavedFileByName(string docName)
        {
            byte[] fileByteArray = { };

            using (ExcelPackage excelPackage = GetSavedExcelPackageByName(docName))
            {
                ExcelWorksheets sheetList = excelPackage.Workbook.Worksheets;

                foreach(ExcelWorksheet sheet in sheetList)
                {
                    sheet.Protection.SetPassword("bimar123");
                    sheet.Protection.AllowEditObject = false;
                    sheet.Protection.AllowEditScenarios = false;
                    sheet.Protection.AllowDeleteColumns = false;
                    sheet.Protection.AllowDeleteRows = false;
                    sheet.Protection.AllowFormatCells = false;
                    sheet.Protection.AllowFormatColumns = false;
                    sheet.Protection.AllowFormatRows = false;
                    sheet.Protection.AllowInsertColumns = false;
                    sheet.Protection.AllowInsertHyperlinks = false;
                    sheet.Protection.AllowInsertRows = false;
                    sheet.Protection.AllowPivotTables = false;
                    sheet.Protection.AllowSelectLockedCells = false;
                    sheet.Protection.AllowSelectUnlockedCells = false;
                    sheet.Protection.AllowSort = false;
                }

                fileByteArray = excelPackage.GetAsByteArray();
            }

            string file = Convert.ToBase64String(fileByteArray);
            return "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + file;
        }
    }
}