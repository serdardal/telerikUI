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

        private UnlockResponseModel OptimizedFindUnlockedCells(CellUnlockModel model)
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
            List<UnlockedTableModel> notNullCells = new List<UnlockedTableModel>();
            List<UnlockedTableModel> shipParticularCells = new List<UnlockedTableModel>();
            List<MergeTableModel> mergedTables = new List<MergeTableModel>();
            List<UnlockedTableModel> mergedDataCells = new List<UnlockedTableModel>();

            FileInfo fi = new FileInfo(path);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheets worksheetList = excelPackage.Workbook.Worksheets;


                //data ve unlock celleri bulur.
                for (int k = 0; k < worksheetList.Count; k++)
                {
                    var currentWorksheet = worksheetList[k];

                    dataCells.Add(new UnlockedTableModel { TableIndex = k, CellList = new List<FilledCellModel>() });
                    notNullCells.Add(new UnlockedTableModel { TableIndex = k, CellList = new List<FilledCellModel>() });
                    shipParticularCells.Add(new UnlockedTableModel { TableIndex = k, CellList = new List<FilledCellModel>() });
                    mergedTables.Add(new MergeTableModel { TableIndex = k, MergedCellList = new List<string>() });
                    mergedDataCells.Add(new UnlockedTableModel { TableIndex = k, CellList = new List<FilledCellModel>() });

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
                                        mergedDataCells[k].CellList.Add(new FilledCellModel { RowIndex = i, ColumnIndex = j, Value = value == null ? null : value.ToString(), Format = format });
                                        mergedTables[k].MergedCellList.Add(mergeAdress);
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

            return new UnlockResponseModel { DataCells = dataCells, NotNullCells = notNullCells, ShipParticularCells = shipParticularCells, MergedTables = mergedTables, MergedDataCells = mergedDataCells };
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
            List<UnlockedTableModel> formulaCells = new List<UnlockedTableModel>();
            List<MergeTableModel> mergedTables = new List<MergeTableModel>();

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
                    formulaCells.Add(new UnlockedTableModel { TableIndex = k, CellList = new List<FilledCellModel>() });
                    mergedTables.Add(new MergeTableModel { TableIndex = k, MergedCellList = new List<string>() });

                    List<string> mergedCellList = currentWorksheet.MergedCells.ToList();
                    //kilitli olmayan ve merge edilmemiş hücreleri bulur ve listeye ekler
                    for (int i = 1; i < 300; i++)
                    {
                        for (int j = 1; j < 300; j++)
                        {
                            var currentCell = currentWorksheet.Cells[i, j];
                            bool locked = currentCell.Style.Locked;
                            bool merged = currentCell.Merge;
                            string formula = currentCell.Formula;

                            if (formula != "")
                            {
                                var value = currentCell.Value;
                                string format = currentCell.Style.Numberformat.Format;
                                formulaCells[k].CellList.Add(new FilledCellModel { RowIndex = i, ColumnIndex = j, Value = value == null ? null : value.ToString(), Format = format });
                            }

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
                                        mergedTables[k].MergedCellList.Add(mergeAdress);
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

            return new UnlockResponseModel { DataCells = dataCells, OnlyUnlockCells = onlyUnlockedCells, NotNullCells = notNullCells, ShipParticularCells=shipParticularCells, FormulaCells=formulaCells, MergedTables=mergedTables };
        }

        [HttpPost]
        public IActionResult GetUnlockedCells([FromBody] CellUnlockModel model)
        {
            UnlockResponseModel response = OptimizedFindUnlockedCells(model);
            return Ok(response);
        }

        public ActionResult SaveFileToTemp(string contentType, string base64, string fileName)
        {
            System.IO.Directory.CreateDirectory(Directory.GetCurrentDirectory() + "\\Temp");
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

            List<CellRecord> DBCellRecords = _excelService.GetCellRecordsByDocName(docName);

            UnlockResponseModel unlockResponseModel = FindUnlockedCells(new CellUnlockModel { DocumentName = docName, IsTemplate = false });
            List<UnlockedTableModel> dataTables = unlockResponseModel.DataCells;
            List<UnlockedTableModel> formulaTables = unlockResponseModel.FormulaCells;
            foreach (UnlockedTableModel dataTable in dataTables)
            {
                UnlockedTableModel formulaTable = formulaTables[dataTable.TableIndex];

                dataTable.CellList.AddRange(formulaTable.CellList);
            }

            List<CellRecord> newCellRecords = new List<CellRecord>();
            List<CellRecord> updatedCellRecords = new List<CellRecord>();
            List<CellRecord> deletedCellRecords = new List<CellRecord>();

            string tempFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Temp", docName + ".xlsx");
            FileInfo fi = new FileInfo(tempFilePath);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheets worksheetList = excelPackage.Workbook.Worksheets;

                foreach(CellRecord cellRecord in DBCellRecords)
                {
                    ExcelWorksheet tempWorksheet = worksheetList[cellRecord.TableIndex];
                    var tempCell = tempWorksheet.Cells[cellRecord.RowIndex, cellRecord.ColumnIndex];
                    //yeni değer nullse silinmiş
                    if (tempCell.Value == null)
                    {
                        deletedCellRecords.Add(cellRecord);
                        
                    }
                    // değer değiştiyse update edilmiş
                    else if (tempCell.Value.ToString() != cellRecord.Data)
                    {
                        updatedCellRecords.Add(new CellRecord
                        {
                            RowIndex = cellRecord.RowIndex,
                            ColumnIndex = cellRecord.ColumnIndex,
                            Data = tempCell.Value.ToString(),
                            TableIndex = cellRecord.TableIndex,
                            TemplateName = templateName,
                            FileName = docName,
                            Date = date
                        });

                    }

                    // kayıtlarda olanları listeden çıkarır,geriye yeni eklenme ihtimali olanlar kalır
                    dataTables[cellRecord.TableIndex].CellList.RemoveAll(x => x.RowIndex == cellRecord.RowIndex && x.ColumnIndex == cellRecord.ColumnIndex);
                }

                foreach(UnlockedTableModel table in dataTables)
                {
                    ExcelWorksheet tempWorksheet = worksheetList[table.TableIndex];

                    List<FilledCellModel> cellList = table.CellList;
                    foreach(FilledCellModel cell in cellList)
                    {
                        var value = tempWorksheet.Cells[cell.RowIndex, cell.ColumnIndex].Value;
                        if (value != null)
                        {
                            newCellRecords.Add(new CellRecord
                            {
                                RowIndex = cell.RowIndex,
                                ColumnIndex = cell.ColumnIndex,
                                Data = value.ToString(),
                                TableIndex = table.TableIndex,
                                TemplateName = templateName,
                                FileName = docName,
                                Date = date
                            }); ;
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
            UnlockResponseModel unlockResponseModel = FindUnlockedCells(new CellUnlockModel { DocumentName = templateName, IsTemplate = true });
            List<UnlockedTableModel> dataTables = unlockResponseModel.DataCells;
            List<UnlockedTableModel> formulaTables = unlockResponseModel.FormulaCells;

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

                foreach(UnlockedTableModel formulaTable in formulaTables)
                {
                    ExcelWorksheet tempWorksheet = worksheetList[formulaTable.TableIndex];
                    List<FilledCellModel> formulaCellList = formulaTable.CellList;

                    foreach (FilledCellModel cell in formulaCellList)
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
                                TableIndex = formulaTable.TableIndex,
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
            RemoveExcelShapesFromExcelPackage(excelPackage);

            ExcelWorkbook excelWorkBook = excelPackage.Workbook;

            foreach (CellRecord cell in cells)
            {
                ExcelWorksheet worksheet = excelWorkBook.Worksheets[cell.TableIndex];
                bool isFormulaCell = worksheet.Cells[cell.RowIndex, cell.ColumnIndex].Formula == "" ? false : true;
                if (!isFormulaCell)
                {
                    ExcelRange range = worksheet.Cells[cell.RowIndex, cell.ColumnIndex];
                    if (range.Style.Numberformat.Format.StartsWith("0"))
                    {
                        range.Value = Int32.Parse(cell.Data);
                    }
                    else if (range.Style.Numberformat.Format.StartsWith("mm"))
                    {
                        range.Value = DateTime.Parse(cell.Data);
                    }
                    else
                    {
                        range.Value = cell.Data;
                    }
                    

                }
            }

            return excelPackage;
        }

        private ExcelPackage GetExcelPackageByTeplateName(string templateName)
        {
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Forms", templateName);
            FileInfo fi = new FileInfo(templatePath);
            ExcelPackage excelPackage = new ExcelPackage(fi);
            RemoveExcelShapesFromExcelPackage(excelPackage);

            return excelPackage;
        }

        [HttpGet("SecondPage/GetProtectedSavedFileByName/{docName}")]
        public string GetProtectedSavedFileByName(string docName)
        {
            byte[] fileByteArray = { };

            using (ExcelPackage excelPackage = GetSavedExcelPackageWithShapesByName(docName))
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

        private ExcelPackage GetSavedExcelPackageWithShapesByName(string docName)
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
                bool isFormulaCell = worksheet.Cells[cell.RowIndex, cell.ColumnIndex].Formula == "" ? false : true;
                if (!isFormulaCell)
                {
                    worksheet.Cells[cell.RowIndex, cell.ColumnIndex].Value = cell.Data;

                }
            }

            return excelPackage;
        }

        private void RemoveExcelShapesFromExcelPackage(ExcelPackage excelPackage)
        {
            ExcelWorkbook excelWorkBook = excelPackage.Workbook;
            ExcelWorksheets excelWorksheets = excelWorkBook.Worksheets;

            foreach(ExcelWorksheet worksheet in excelWorksheets)
            {
                OfficeOpenXml.Drawing.ExcelDrawings drawings = worksheet.Drawings;

                List<OfficeOpenXml.Drawing.ExcelDrawing> drawingRemoveList = new List<OfficeOpenXml.Drawing.ExcelDrawing>();

                foreach(OfficeOpenXml.Drawing.ExcelDrawing drawing in drawings)
                {
                    if(drawing.GetType() == typeof (OfficeOpenXml.Drawing.ExcelShape)){
                        drawingRemoveList.Add(drawing);
                    }
                }

                foreach(OfficeOpenXml.Drawing.ExcelDrawing drawingToRemove in drawingRemoveList)
                {
                    drawings.Remove(drawingToRemove);
                }
            }
        }
    }
}