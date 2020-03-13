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

namespace Deneme.Controllers
{
    public class HomeController : Controller
    {
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
            List<UnlockedCellModel> unlockedCells = new List<UnlockedCellModel>();
            
            FileInfo fi = new FileInfo(path);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[0];

                for (int i = 1; i < 64; i++)
                {
                    for (int j = 1; j < 11; j++)
                    {
                        bool locked = firstWorksheet.Cells[i, j].Style.Locked;

                        if (!locked) unlockedCells.Add(new UnlockedCellModel { RowIndex = i, ColumnIndex = j });
                    }
                }


            }

            return Ok(unlockedCells);
        }

        [HttpPost]
        public IActionResult WriteToExcelFile([FromBody] FileSaveModel model)
        {
            List<FilledCellModel> cells = model.CellList;
            string path = Path.Combine(Directory.GetCurrentDirectory(), "Forms", model.TemplateName);
            FileInfo fi = new FileInfo(path);
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[0];

                for (int i = 0; i < cells.Count; i++)
                {
                    firstWorksheet.Cells[cells[i].RowIndex, cells[i].ColumnIndex].Value = cells[i].Value;
                }

                string savePath = Path.Combine(Directory.GetCurrentDirectory(),
                    "Saves", 
                    model.TemplateName.Replace(".xlsx","")+"_"+model.DocumentName+"_"+model.Date.ToString("dd.MM.yyyy") + ".xlsx");
                FileInfo savePathFI = new FileInfo(savePath);
                excelPackage.SaveAs(savePathFI);

            }

            return Ok();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
