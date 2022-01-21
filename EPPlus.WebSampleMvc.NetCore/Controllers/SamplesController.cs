using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using EPPlus.WebSampleMvc.NetCore.Samples.Sample1;
using EPPlus.WebSampleMvc.NetCore.Samples.Sample2;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;

namespace EPPlus.WebSampleMvc.NetCore.Controllers
{
    public class SamplesController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        private const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        /// <summary>
        /// Simple example on how to create a workbook without
        /// access to the file system and return it as a file.
        /// </summary>
        /// <returns>A workbook as a file with filename and content type set</returns>
        [HttpGet]
        public IActionResult WebSample1()
        {
            using(var package = new ExcelPackage())
            {
                Sample1.InitWorkbook(package);
                var excelData = package.GetAsByteArray();
                var fileName = "MyWorkbook.xlsx";
                return File(excelData, ContentType, fileName);
            }
        }

        /// <summary>
        /// This sample uses the LoadFromCollection method of EPPlus to load data into a worksheet.
        /// </summary>
        /// <returns>A workbook as a file with filename and content type set</returns>
        /// <seealso cref="Sample2"/>
        [HttpGet]
        public IActionResult WebSample2()
        {
            using (var package = new ExcelPackage())
            {
                Sample2.InitWorkbook(package);
                return File(package.GetAsByteArray(), ContentType, "Sample2.xlsx");
            }
        }
    }
}