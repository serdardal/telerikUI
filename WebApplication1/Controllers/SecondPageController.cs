using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Deneme.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using WebApplication1.Models;
using WebApplication1.Models.Contracts;
using WebApplication1.Models.DbModels;
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
            IndexModel model = new IndexModel { OpenInNewTab = false };
            return View(model);
        }

        [HttpGet]
        public IActionResult GetTemplateNames()
        {
            //Forms klasörü altındaki şablon isimlerini döndürür.
            string[] excelFiles = Directory.GetFiles(Path.Combine(Directory.GetCurrentDirectory(), "Forms"), "*.xlsx")
                                     .Select(Path.GetFileName)
                                     .ToArray();

            return Ok(excelFiles);
        }

        [HttpPost]
        public GetTemplateResponse GetTemplateByName([FromBody] OpenTemplateRequestModel requestModel)
        {
            string templateName = requestModel.TemplateName;
            //ship particular değişkenlerin bulunduğu dictionary
            Dictionary<string, string> variableDictionary = new Dictionary<string, string>();

            //databasede kayıtlı değişkenlerin dictionary'e eklenmesi
            List<Default> defaults = _excelService.GetDefaults();
            foreach(Default item in defaults)
            {
                variableDictionary.Add(item.Key, item.Value);
            }


            byte[] fileByteArray = { };
            CellGroupModel _cellGroupModel;
            //gönderilmeden önce ship particular değişkenlerin doldurulması
            using (ExcelPackage excelPackage = GetExcelPackageByTeplateName(templateName))
            {
                CellGroupModel cellGroups = FindCellGroups(excelPackage, templateName);
                _cellGroupModel = cellGroups;

                List<TableModel> shipParticularCellTables = cellGroups.ShipParticularCellTables;

                ExcelWorksheets worksheetList = excelPackage.Workbook.Worksheets;

                if (!string.IsNullOrEmpty(requestModel.LogoName))
                {
                    ChangePicture(excelPackage.Workbook, requestModel.LogoName);
                }

                for (int i = 0; i < shipParticularCellTables.Count; i++)
                {
                    ExcelWorksheet templateWorksheet = worksheetList[i];
                    TableModel table = shipParticularCellTables[i];
                    List<CellModel> cellList = table.CellList;

                    foreach (CellModel cell in cellList)
                    {
                        var value = templateWorksheet.Cells[cell.RowIndex, cell.ColumnIndex].Value;
                        if (value != null) {
                            string key = value.ToString();
                            if (variableDictionary.ContainsKey(key))
                            {
                                templateWorksheet.Cells[cell.RowIndex, cell.ColumnIndex].Value = variableDictionary[key];
                            }
                        }
                        
                    }
                }

                RemoveExcelShapesFromExcelPackage(excelPackage);
                RemoveEndMarksFrowWorkBook(excelPackage.Workbook, templateName);
                fileByteArray = excelPackage.GetAsByteArray();
            }

            string file = Convert.ToBase64String(fileByteArray);
            string base64File = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + file;

            return new GetTemplateResponse {
                Base64File = base64File,
                NotMergedDataCellTables = _cellGroupModel.NotMergedDataCellTables,
                NotNullCellTables = _cellGroupModel.NotNullCellTables,
                ShipParticularCellTables = _cellGroupModel.ShipParticularCellTables,
                MergedRangesTables = _cellGroupModel.MergedRangesTables,
                EndMarks = _cellGroupModel.EndMarks,
                CustomFormattedCellTables = _cellGroupModel.CustomFormattedCellTables
            };
        }

        [HttpGet]
        public IActionResult GetSavedFileNamesFromDB()
        {
            //databasedeki kayıtlı dosya isimlerini liste şeklinde döndürür.
            List<string> fileNames = _excelService.GetSavedFileNames();
            return Ok(fileNames);
        }

        [HttpGet("SecondPage/GetSavedFileByName/{fileName}")]
        public GetSavedFilesResponse GetSavedFileByName(string fileName)
        {
            byte[] fileByteArray = { };
            CellGroupModel _cellGroupModel;

            var templateName = FindTemplateNameFromFileName(fileName);

            using (ExcelPackage excelPackage = GetSavedExcelPackageByName(fileName))
            {
                CellGroupModel cellGroups = FindCellGroups(excelPackage, templateName);
                _cellGroupModel = cellGroups;

                // formun logosu var ise var olan logo ile değiştir
                string logo = _excelService.GetLogoByName(fileName);
                if (!string.IsNullOrEmpty(logo))
                {
                    ChangePicture(excelPackage.Workbook, logo);
                }

                RemoveExcelShapesFromExcelPackage(excelPackage);
                RemoveEndMarksFrowWorkBook(excelPackage.Workbook, templateName);

                fileByteArray = excelPackage.GetAsByteArray();
            }

            string file = Convert.ToBase64String(fileByteArray);
            string base64File = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64," + file;

            return new GetSavedFilesResponse
            {
                Base64File = base64File,
                NotMergedDataCellTables = _cellGroupModel.NotMergedDataCellTables,
                NotNullCellTables = _cellGroupModel.NotNullCellTables,
                ShipParticularCellTables = _cellGroupModel.ShipParticularCellTables,
                MergedRangesTables = _cellGroupModel.MergedRangesTables,
                EndMarks = _cellGroupModel.EndMarks,
                CustomFormattedCellTables = _cellGroupModel.CustomFormattedCellTables
            };
        }

        [HttpPost]
        public string GetProtectedSavedFileByName([FromBody] ExportRequestModel exportModel)
        {
            string fileName = exportModel.FileName;
            List<ColoredCellModel> coloredCellList = exportModel.ColoredCellList;

            //protected olarak export etmede kullanılır.
            byte[] fileByteArray = { };

            using (ExcelPackage excelPackage = GetSavedExcelPackageWithShapesByName(fileName))
            {
                // formun logosu var ise var olan logo ile değiştir
                string logo = _excelService.GetLogoByName(fileName);
                if (!string.IsNullOrEmpty(logo))
                {
                    ChangePicture(excelPackage.Workbook, logo);
                }

                //zemin rengi değiştirilecek hücrelerin işlenmesi
                ColorCells(excelPackage.Workbook, coloredCellList);

                ExcelWorksheets sheetList = excelPackage.Workbook.Worksheets;

                //sheetler için protect ayarları
                foreach (ExcelWorksheet sheet in sheetList)
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

        [HttpGet("SecondPage/OpenFileReadonlyInNewTab/{fileName}/{readOnly}")]
        public IActionResult OpenFileReadonlyInNewTab(string fileName, bool readOnly) {
            IndexModel model = new IndexModel { OpenInNewTab = true, FileName = fileName, ReadOnly = readOnly };
            return View("Index", model);
        }

        [HttpPost]
        public ActionResult SaveFileToTemp([FromBody] SaveFileToTempRequestModel requestModel)
        {
            string base64 = requestModel.Base64;
            string fileName = requestModel.FileName;
            string logoName = requestModel.LogoName;
            //dosya kayıt edilirse veya update edilirse Temp klasörü altına kaydedilir.

            //Temp klasörü yok ise oluştur.
            System.IO.Directory.CreateDirectory(Directory.GetCurrentDirectory() + "\\Temp");
            var fileContents = Convert.FromBase64String(base64);
            System.IO.File.WriteAllBytes(Directory.GetCurrentDirectory() + $"\\Temp\\{fileName}.xlsx", fileContents);

            //dosya kaydedildikten sonra üzerindeki veriler database ile senkronize edilir.
            SyncDataWithDB(fileName, logoName);

            IndexModel model = new IndexModel { OpenInNewTab = false };

            return View("Index", model);
        }

        private CellGroupModel FindCellGroups(ExcelPackage excelPackage, string templateName)
        {
            ExcelWorksheets worksheetList = excelPackage.Workbook.Worksheets;

            List<TableModel> dataCells = new List<TableModel>();
            List<TableModel> notMergedDataCells = new List<TableModel>();
            List<MergeTableModel> mergedTables = new List<MergeTableModel>();
            List<NotNullTableModel> notNullCells = new List<NotNullTableModel>();
            List<TableModel> shipParticularCells = new List<TableModel>();
            List<CustomFormattedTableModel> customFormattedCells = new List<CustomFormattedTableModel>();
            List<TableModel> formulaCellTables = new List<TableModel>();

            List<EndMark> endMarks = new List<EndMark>();
            //endmarkların bulunması
            var endMarkRecords = _excelService.GetEndMarksofTemplate(templateName);
            if (endMarkRecords.Count == 0)
            {
                endMarkRecords = FindEndMarksInTemplate(templateName);
            }
            endMarks = endMarkRecords;

            //sheetlerin gezilmesi
            for (int k = 0; k < worksheetList.Count; k++)
            {
                var currentWorksheet = worksheetList[k];

                dataCells.Add(new TableModel { TableIndex = k, CellList = new List<CellModel>() });
                notMergedDataCells.Add(new TableModel { TableIndex = k, CellList = new List<CellModel>() });
                mergedTables.Add(new MergeTableModel { TableIndex = k, MergedCellList = new List<string>() });
                notNullCells.Add(new NotNullTableModel { TableIndex = k, CellList = new List<CellModelWithValue>() });
                shipParticularCells.Add(new TableModel { TableIndex = k, CellList = new List<CellModel>() });
                customFormattedCells.Add(new CustomFormattedTableModel { TableIndex = k, CellList = new List<CellModelWithFormat>() });
                formulaCellTables.Add(new TableModel { TableIndex = k, CellList = new List<CellModel>() });

                List<string> mergedCellList = currentWorksheet.MergedCells.ToList();

                //bir sheet için {E} sınır belirlenemez ise hücreler 300x300 bir alanda aranır.
                int countOfRowsToSearch = 300;
                int countOfColumnsToSearch = 300;

                //aranacak sınırın belirlenmesi
                if (endMarkRecords.Count > 0)
                {
                    foreach (EndMark endMark in endMarkRecords)
                    {
                        if (endMark.SheetIndex == k)
                        {
                            countOfRowsToSearch = endMark.RowIndex;
                            countOfColumnsToSearch = endMark.ColumnIndex;
                        }
                    }
                }

                //kilitli olmayan ve merge edilmemiş hücreleri bulur ve listeye ekler
                for (int i = 1; i < countOfRowsToSearch; i++) //satır
                {
                    for (int j = 1; j < countOfColumnsToSearch; j++) //sütun
                    {
                        var currentCell = currentWorksheet.Cells[i, j];
                        bool locked = currentCell.Style.Locked;
                        bool merged = currentCell.Merge;


                        if (!locked)
                        {
                            var value = currentCell.Value;
                            string format = currentCell.Style.Numberformat.Format;
                            var newCell = new CellModel { RowIndex = i, ColumnIndex = j};

                            if (!merged)
                            {
                                notMergedDataCells[k].CellList.Add(newCell);
                                dataCells[k].CellList.Add(newCell);
                            }
                            else
                            {
                                if(isMergedDataCell(currentWorksheet, i, j))
                                {
                                    //verinin tutulduğu hücre olduğu için dataCells e eklenir.
                                    dataCells[k].CellList.Add(newCell);

                                    //range olarak açılması gereken adres
                                    var mergeAdress = currentWorksheet.MergedCells[i, j];
                                    mergedTables[k].MergedCellList.Add(mergeAdress);
                                }
                            }

                            if (isNotNullCell(currentCell))
                            {
                                var newNotNullCell = new CellModelWithValue { RowIndex = i, ColumnIndex = j, Value= value == null? null : value.ToString() };
                                notNullCells[k].CellList.Add(newNotNullCell);
                            }

                            if (isShipParticularCell(currentCell)) shipParticularCells[k].CellList.Add(newCell);

                            if (isCustomFormattedCell(currentCell))
                            {
                                var newCustomFormattedCell = new CellModelWithFormat { RowIndex = i, ColumnIndex = j, Format = format };
                                customFormattedCells[k].CellList.Add(newCustomFormattedCell);
                            }

                        }
                        else // formül hücreleri locked durumdadır.
                        {
                            if (isFormulaCell(currentCell)) {
                                var value = currentCell.Value;
                                string format = currentCell.Style.Numberformat.Format;
                                var newCell = new CellModel { RowIndex = i, ColumnIndex = j};

                                formulaCellTables[k].CellList.Add(newCell);
                            } 
                        }
                    }
                }

            }

            return new CellGroupModel
            {
                DataCellTables = dataCells,
                NotMergedDataCellTables = notMergedDataCells,
                NotNullCellTables = notNullCells,
                ShipParticularCellTables = shipParticularCells,
                MergedRangesTables = mergedTables,
                EndMarks = endMarks,
                CustomFormattedCellTables = customFormattedCells,
                FormulaCellTables = formulaCellTables
            };
        }

        private bool isNotNullCell(ExcelRange range)
        {
            var value = range.Value;
            if (value != null && value.ToString().StartsWith("{NN;"))
            {
                return true;
            }
            return false;
        }

        private bool isShipParticularCell(ExcelRange range)
        {
            var value = range.Value;
            if (value != null && !value.ToString().StartsWith("{NN;") && value.ToString().StartsWith("{") && value.ToString().EndsWith("}"))
            {
                return true;
            }
            return false;
        }

        private bool isCustomFormattedCell(ExcelRange range)
        {
            string format = range.Style.Numberformat.Format;

            if (format.StartsWith("["))
            {
                return true;
            }
            return false;
        }

        private bool isMergedDataCell(ExcelWorksheet currentWorksheet, int row, int column)
        {
            //merge hücrelerden sol üstte olan data celldir.
            var mergeAdress = currentWorksheet.MergedCells[row, column];

            //master = veriyi tutan sol üstteki hücre
            string masterCellName = mergeAdress.Split(":")[0];
            var masterCell = currentWorksheet.Cells[masterCellName];

            if (masterCell.Start.Row == row && masterCell.Start.Column == column) 
            {
                return true;
            }

            return false;
        }

        private bool isFormulaCell(ExcelRange range)
        {
            string formula = range.Formula;

            // formül içeriyorsa formül celleri listesine eklenir.
            if (formula != "")
            {
                return true;
            }

            return false;
        }

        private void SyncDataWithDB(string fileName, string logoName)
        {
            //kayıt edilen dosyanın yeni bir dosya mı yoksa kayıtlı bir dosya mı olduğunun belirlenmesi
            var cells = _excelService.GetCellRecordsByDocName(fileName);
            if(cells.Count > 0)// kayıt bulunuyor yani update işlemi
            {
                UpdateExistingFileInDB(fileName);
            }
            else // kayıt yok yani ekleme işlemi
            {
                AddNewRecordsToDB(fileName, logoName);
            }
        }

        private void UpdateExistingFileInDB(string fileName)
        {
            string templateName = FindTemplateNameFromFileName(fileName);
            DateTime date = FindDateFromFileName(fileName);

            List<CellRecord> DBCellRecords = _excelService.GetCellRecordsByDocName(fileName);

            CellGroupModel _cellGroups;
            using (ExcelPackage excelPackage = GetExcelPackageByTeplateName(templateName))
            {
                CellGroupModel cellGroups = FindCellGroups(excelPackage, templateName);
                _cellGroups = cellGroups;
            }

            //değişiklikler data cellerde veya formül cellerinde olabilir şablon üzerinden bu hücrelerin alınması.
            List<TableModel> dataTables = _cellGroups.DataCellTables;
            List<TableModel> formulaTables = _cellGroups.FormulaCellTables;
            //bakılacak cellerin bir yerde toplanması.
            foreach (TableModel dataTable in dataTables)
            {
                TableModel formulaTable = formulaTables[dataTable.TableIndex];

                dataTable.CellList.AddRange(formulaTable.CellList);
            }

            //değişiklikler şu şekilde olabilir:
            //yeni kayıtlar eklenmiş olabilir.
            //var olan kayıtların değeri değişmiş olabilir.
            //var olan kayıt silinmiş olabilir.
            List<CellRecord> newCellRecords = new List<CellRecord>();
            List<CellRecord> updatedCellRecords = new List<CellRecord>();
            List<CellRecord> deletedCellRecords = new List<CellRecord>();

            using (ExcelPackage excelPackage = GetTemporaryExcelPackageByName(fileName + ".xlsx"))
            {
                ExcelWorksheets worksheetList = excelPackage.Workbook.Worksheets;

                // güncellenmiş veya silinmiş kayıtların belirlenmesi.
                foreach(CellRecord cellRecord in DBCellRecords)
                {
                    var cell = dataTables[cellRecord.TableIndex].CellList.FirstOrDefault(x => x.RowIndex == cellRecord.RowIndex && x.ColumnIndex == cellRecord.ColumnIndex);
                    ExcelWorksheet tempWorksheet = worksheetList[cellRecord.TableIndex];
                    var tempCell = tempWorksheet.Cells[cellRecord.RowIndex, cellRecord.ColumnIndex];
                    //yeni değer nullse silinmiş
                    if (tempCell.Value == null)
                    {
                        deletedCellRecords.Add(cellRecord);
                        
                    }
                    // değer değiştiyse update edilmiş
                    else if (tempCell.Text != cellRecord.Data || tempCell.Value.ToString() != cellRecord.Data)
                    {
                        string value = tempCell.Text;
                        string type = null;
                        string format = tempCell.Style.Numberformat.Format;
                        if (format != null)
                        {
                            type = FindTypeOfCell(format);
                        }
                        if (type == "number") value = tempCell.Value.ToString();

                        cellRecord.Data = value;
                        updatedCellRecords.Add(cellRecord);
                    }

                    // kayıtlarda olanları listeden çıkarır,geriye yeni eklenme ihtimali olanlar kalır
                    dataTables[cellRecord.TableIndex].CellList.Remove(cell);
                }

                //yeni eklenmiş kayıtların belirlenmesi.
                foreach(TableModel table in dataTables)
                {
                    ExcelWorksheet tempWorksheet = worksheetList[table.TableIndex];

                    List<CellModel> cellList = table.CellList;
                    foreach(CellModel cell in cellList)
                    {
                        var tempCell = tempWorksheet.Cells[cell.RowIndex, cell.ColumnIndex];
                        if (tempCell.Value != null)
                        {
                            string value = tempCell.Text;
                            string type = null;
                            string format = tempCell.Style.Numberformat.Format;
                            if (format != null)
                            {
                                type = FindTypeOfCell(format);
                            }
                            if (type == "number") value = tempCell.Value.ToString();

                            newCellRecords.Add(new CellRecord
                            {
                                RowIndex = cell.RowIndex,
                                ColumnIndex = cell.ColumnIndex,
                                Data = value,
                                TableIndex = table.TableIndex,
                                TemplateName = templateName,
                                FileName = fileName,
                                Date = date,
                                Type = type
                            }); ;
                        }
                    }
                }

            }

            _excelService.UpdateCells(newCellRecords, updatedCellRecords, deletedCellRecords);
        }

        private void AddNewRecordsToDB(string fileName, string logoName)
        {
            string templateName = FindTemplateNameFromFileName(fileName);
            DateTime date = FindDateFromFileName(fileName);

            CellGroupModel _cellGroups;
            using (ExcelPackage excelPackage = GetExcelPackageByTeplateName(templateName))
            {
                CellGroupModel cellGroups = FindCellGroups(excelPackage, templateName);
                _cellGroups = cellGroups;
            }

            //değişiklikler data cellerde veya formül cellerinde olabilir şablon üzerinden bu hücrelerin alınması.
            List<TableModel> dataTables = _cellGroups.DataCellTables;
            List<TableModel> formulaTables = _cellGroups.FormulaCellTables;

            //eklenecek yeni kayıtlar listesi
            List<CellRecord> newCellRecords = new List<CellRecord>();

            using (ExcelPackage excelPackage = GetTemporaryExcelPackageByName(fileName + ".xlsx"))
            {
                ExcelWorksheets worksheetList = excelPackage.Workbook.Worksheets;

                //data cellerin satır ve sütunlarını biliyoruz
                //temp dosya üzerinde bu koordinatlara gidilir ve null değilse değer alınır.
                foreach(TableModel table in dataTables)
                {
                    ExcelWorksheet tempWorksheet = worksheetList[table.TableIndex];

                    List<CellModel> cellList = table.CellList;

                    foreach(CellModel cell in cellList)
                    {
                        var tempCell = tempWorksheet.Cells[cell.RowIndex, cell.ColumnIndex];
                        
                        if (tempCell.Value != null)
                        {
                            string value = tempCell.Text;
                            string type = null;
                            string format = tempCell.Style.Numberformat.Format;
                            if (format != null)
                            {
                                type = FindTypeOfCell(format);
                            }
                            if (type == "number") value = tempCell.Value.ToString();

                            newCellRecords.Add(new CellRecord
                            {
                                RowIndex = cell.RowIndex,
                                ColumnIndex = cell.ColumnIndex,
                                Data = value,
                                TableIndex = table.TableIndex,
                                TemplateName = templateName,
                                FileName = fileName,
                                Date = date,
                                Type = type,
                                Logo = logoName,
                            });
                        }
                    }
                }

                //formül cellerin satır ve sütunlarını biliyoruz
                //temp dosya üzerinde bu koordinatlara gidilir ve null değilse değer alınır.
                foreach (TableModel formulaTable in formulaTables)
                {
                    ExcelWorksheet tempWorksheet = worksheetList[formulaTable.TableIndex];
                    List<CellModel> formulaCellList = formulaTable.CellList;

                    foreach (CellModel cell in formulaCellList)
                    {
                        var tempCell = tempWorksheet.Cells[cell.RowIndex, cell.ColumnIndex];

                        if (tempCell.Value != null)
                        {
                            string value = tempCell.Text;
                            string type = null;
                            string format = tempCell.Style.Numberformat.Format;
                            if (format != null)
                            {
                                type = FindTypeOfCell(format);
                            }
                            if (type == "number") value = tempCell.Value.ToString();

                            newCellRecords.Add(new CellRecord
                            {
                                RowIndex = cell.RowIndex,
                                ColumnIndex = cell.ColumnIndex,
                                Data = value,
                                TableIndex = formulaTable.TableIndex,
                                TemplateName = templateName,
                                FileName = fileName,
                                Date = date,
                                Type = type
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

            //kayıtlı dosya isimleri şablon ismi ile başlıyor mu kontrol edilir.
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
            //kayıtlı dosyaların son 10 karakteri tarihi içeriyor
            return DateTime.Parse(fileName.Substring(fileName.Length - 10, 10));
        }

        private ExcelPackage GetSavedExcelPackageByName(string fileName)
        {
            string templateName = _excelService.GetTemplateName(fileName);
            if(templateName == null)
            {
                templateName = FindTemplateNameFromFileName(fileName);
            }
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Forms", templateName);

            //databasede kayıtlı hücreler
            List<CellRecord> cells = _excelService.GetCellRecordsByDocName(fileName);

            FileInfo fi = new FileInfo(templatePath);
            ExcelPackage excelPackage = new ExcelPackage(fi);

            ExcelWorkbook excelWorkBook = excelPackage.Workbook;

            //kayıtlı hücrelerin şablon içerisine doldurulması
            foreach (CellRecord cell in cells)
            {
                ExcelWorksheet worksheet = excelWorkBook.Worksheets[cell.TableIndex];
                bool isFormulaCell = worksheet.Cells[cell.RowIndex, cell.ColumnIndex].Formula == "" ? false : true;
                if (!isFormulaCell)
                {
                    ExcelRange range = worksheet.Cells[cell.RowIndex, cell.ColumnIndex];
                    try
                    {
                        string type = FindTypeOfCell(range.Style.Numberformat.Format);

                        if(type == "text")
                        {
                            range.Value = cell.Data;
                        }
                        else if (type == "number")
                        {
                            range.Value = double.Parse(cell.Data);
                        }
                        else if (type == "date")
                        {
                            range.Value = DateTime.Parse(cell.Data);
                        }
                        else if (type == "time")
                        {
                            range.Value = DateTime.Parse(cell.Data);
                        }
                        else
                        {
                            range.Value = cell.Data;
                        }
                    }
                    catch (Exception)
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

            return excelPackage;
        }

        private ExcelPackage GetSavedExcelPackageWithShapesByName(string fileName)
        {
            //eğer excelshape var ise bunun export edilen protected sheetlerde görünmesi istiyoruz

            string templateName = _excelService.GetTemplateName(fileName);
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Forms", templateName);

            List<CellRecord> cells = _excelService.GetCellRecordsByDocName(fileName);

            FileInfo fi = new FileInfo(templatePath);
            ExcelPackage excelPackage = new ExcelPackage(fi);

            ExcelWorkbook excelWorkBook = excelPackage.Workbook;
            RemoveEndMarksFrowWorkBook(excelWorkBook, templateName);

            //şablon databasedeki kayıtlı hücreler ile doldurulur.
            foreach (CellRecord cell in cells)
            {
                ExcelWorksheet worksheet = excelWorkBook.Worksheets[cell.TableIndex];
                bool isFormulaCell = worksheet.Cells[cell.RowIndex, cell.ColumnIndex].Formula == "" ? false : true;
                if (!isFormulaCell)
                {
                    ExcelRange range = worksheet.Cells[cell.RowIndex, cell.ColumnIndex];
                    try
                    {
                        string type = FindTypeOfCell(range.Style.Numberformat.Format);

                        if (type == "text")
                        {
                            range.Value = cell.Data;
                        }
                        else if (type == "number")
                        {
                            range.Value = double.Parse(cell.Data);
                        }
                        else if (type == "date")
                        {
                            range.Value = DateTime.Parse(cell.Data);
                        }
                        else if (type == "time")
                        {
                            range.Value = double.Parse(cell.Data);
                        }
                        else
                        {
                            range.Value = cell.Data;
                        }
                    }
                    catch (Exception)
                    {
                        range.Value = cell.Data;
                    }

                }
            }

            return excelPackage;
        }

        private ExcelPackage GetTemporaryExcelPackageByName(string fileName)
        {
            string path = Path.Combine(Directory.GetCurrentDirectory(), "Temp", fileName);
            FileInfo fi = new FileInfo(path);
            ExcelPackage excelPackage = new ExcelPackage(fi);

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
                    //drawingin tipi ExcelShape ise çıkar
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

        private string FindTypeOfCell(string format)
        {
            //text için format "@"
            if (format.StartsWith("@"))
            {
                return "text";
            }
            //date için format "dd-mm-yy" şeklinde 
            else if (format.StartsWith("m") || format.StartsWith("d") || format.StartsWith("y"))
            {
                return "date";
            }
            //number için [Blue][=1]0; // 0.0 // #.##0 gibi formatlar gelebilir
            else if (format.StartsWith("[") || format.StartsWith("0") ||format.StartsWith("#"))
            {
                return "number";
            }
            //hour için "hh:mm:ss" şeklinde
            else if (format.StartsWith("h"))
            {
                return "time";
            }

            return null;
        }

        private List<EndMark> FindEndMarksInTemplate(string templateName)
        {
            List<EndMark> endMarks = new List<EndMark>();

            string path = Path.Combine(Directory.GetCurrentDirectory(), "Forms", templateName);
            FileInfo fi = new FileInfo(path);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheets worksheetList = excelPackage.Workbook.Worksheets;

                for (int k = 0; k < worksheetList.Count; k++) //sheet index
                {
                    var currentWorksheet = worksheetList[k];
                    bool found = false;

                    for (int j = 1; j < 300; j++)// column index
                    {
                        if (found) break;

                        for (int i = 1; i < 300; i++)//row index
                        {
                            var currentCell = currentWorksheet.Cells[i, j];
                            var value = currentCell.Value;

                            if(value != null && value.ToString() == "{E}")
                            {
                                endMarks.Add(new EndMark { TemplateName = templateName, SheetIndex = k, RowIndex = i, ColumnIndex = j }); ;                                found = true;

                                break;
                            }
                        }
                    }
                }

                _excelService.AddEndMarks(endMarks);

                return endMarks;
            }
        }

        private void RemoveEndMarksFrowWorkBook(ExcelWorkbook workBook, string templateName)
        {
            List<EndMark> endMarks = _excelService.GetEndMarksofTemplate(templateName);

            ExcelWorksheets worksheetList = workBook.Worksheets;

            foreach(EndMark endMark in endMarks)
            {
                ExcelWorksheet worksheet = worksheetList[endMark.SheetIndex];

                worksheet.Cells[endMark.RowIndex, endMark.ColumnIndex].Value = null;
            }


            for (int k = 0; k < worksheetList.Count; k++) //sheet index
            {
                var currentWorksheet = worksheetList[k];

                for (int j = 1; j < 300; j++)// column index
                {

                    for (int i = 1; i < 300; i++)//row index
                    {
                        var currentCell = currentWorksheet.Cells[i, j];
                        var value = currentCell.Value;

                        if (value != null && value.ToString().StartsWith("{NN;"))
                        {
                            currentCell.Value = null;
                        }
                    }
                }
            }

        }

        private void ColorCells(ExcelWorkbook workBook, List<ColoredCellModel> coloredCells)
        {
            ExcelWorksheets excelWorksheets = workBook.Worksheets;

            foreach(ColoredCellModel coloredCell in coloredCells)
            {
                ExcelWorksheet worksheet = excelWorksheets[coloredCell.SheetIndex];

                Color color = ColorTranslator.FromHtml(coloredCell.Color);
                worksheet.Cells[coloredCell.RowIndex, coloredCell.ColumnIndex].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheet.Cells[coloredCell.RowIndex, coloredCell.ColumnIndex].Style.Fill.BackgroundColor.SetColor(color);
            }
        }

        private void ChangePicture(ExcelWorkbook workBook, string logoName)
        {
            ExcelWorksheets excelWorksheets = workBook.Worksheets;

            foreach (ExcelWorksheet worksheet in excelWorksheets)
            {
                List<OfficeOpenXml.Drawing.ExcelDrawing> changeList = new List<OfficeOpenXml.Drawing.ExcelDrawing>();

                //degistirilecek drawinglerin bulunması
                foreach (OfficeOpenXml.Drawing.ExcelDrawing drawing in worksheet.Drawings)
                {
                    if (drawing.Name.StartsWith("Logo"))
                    {
                        changeList.Add(drawing);
                    }
                }

                //drawinglerin imageları değiştiriliyor, geri kalan ayarları değişmemiş oluyor.
                foreach (OfficeOpenXml.Drawing.ExcelPicture changingDrawing in changeList)
                {
                    using (Image newImage = Image.FromFile(Path.Combine(Directory.GetCurrentDirectory(), "Images", logoName)))
                    {
                        changingDrawing.Image = newImage;
                    }
                }
            }
        }
    }
}