using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.Controllers
{
    public class JsonExportController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Sample1()
        {
            return View();
        }

        [HttpGet, Route("/api/currencyrates")]
        public async Task<JsonResult> Sample1GetJsonData()
        {
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Currencies");
                var csvFileInfo = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"data\\currencies2011weekly.csv"));
                var format = new ExcelTextFormat
                {
                    Delimiter = ';',
                    Culture = CultureInfo.InvariantCulture,
                    DataTypes = new eDataTypes[] { eDataTypes.DateTime, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number }                    
                };
                var range = await sheet.Cells["A1"].LoadFromTextAsync(csvFileInfo, format);
                sheet.Cells[range.Start.Row, 1, range.End.Row, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                sheet.Cells[range.Start.Row, 2, range.End.Row, 5].Style.Numberformat.Format = "#,##0.0000";
                var jsonData = range.ToJson(x=>x.AddDataTypesOn = eDataTypeOn.OnColumn);
                return Json(jsonData);
            }
        }
    }
}
