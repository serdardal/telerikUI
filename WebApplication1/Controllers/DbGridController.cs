using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using WebApplication1.Models;
using WebApplication1.Models.DbModels;
using WebApplication1.Services;

namespace WebApplication1.Controllers
{
    public class DbGridController : Controller
    {
        IExcelService _excelService;

        public DbGridController(IExcelService excelService)
        {
            _excelService = excelService;
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult GetAllRecords()
        {
            List<CellRecord> records = _excelService.GetAllRecords();
            return Ok(records);
        }

        public IActionResult ClearAllRecords()
        {
            var clear = _excelService.ClearAllRecords();

            return Ok();
        }
    }
}