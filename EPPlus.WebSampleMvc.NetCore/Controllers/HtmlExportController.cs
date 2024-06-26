﻿using EPPlus.WebSampleMvc.NetCore.Models.HtmlExport;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.Controllers
{
    public class HtmlExportController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        private const string ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        [HttpGet]
        public IActionResult ExportTable1()
        {
            var model = new ExportTable1Model();
            model.SetupSampleData(0);
            model.TableStyle = "Dark1";
            return View(model);
        }

        [HttpPost, ValidateAntiForgeryToken]
        public IActionResult ExportTable1(ExportTable1Model model)
        {
            if(!Enum.TryParse(model.TableStyle, out TableStyles ts))
            {
                ts = TableStyles.Light1;
            }
            ViewData["TableStyle"] = ts.ToString();
            model.SetupSampleData(model.Theme, ts);
            if(model.GetWorkbook)
            {
                return File(model.WorkbookBytes, ContentType, "EPPlusHtmlSample1.xlsx");
            }
            return View(model);
        }

        public IActionResult ExportTable2(string style)
        {
            if (!Enum.TryParse(style, out TableStyles ts))
            {
                ts = TableStyles.Dark1;
            }
            ViewData["TableStyle"] = ts.ToString();
            var model = new ExportTable2Model();
            model.SetupSampleData(ts);
            return View(model);
        }

        public IActionResult ExportTable3(string style)
        {
            if (!Enum.TryParse(style, out TableStyles ts))
            {
                ts = TableStyles.Light2;
            }
            var model = new ExportTable3Model();
            model.SetupSampleData(ts);
            return View(model);
        }

        public IActionResult ExportTable4(string style)
        {
            if (!Enum.TryParse(style, out TableStyles ts))
            {
                ts = TableStyles.Light2;
            }
            var model = new ExportTable4Model();
            model.SetupSampleData(ts);
            return View(model);
        }

        public IActionResult ExportRange5()
        {
            var model = new ExportRange5Model();
            model.SetupSampleData();
            return View(model);
        }

        public IActionResult ExportRanges6()
        {
            var model = new ExportRanges6Model();
            model.SetupSampleData();
            return View(model);
        }

        [HttpGet]
        public IActionResult ExportCf7()
        {
            var model = new ExportCf7Model();
            model.SetupSampleData();
            return View(model);
        }

        [HttpPost, ValidateAntiForgeryToken]
        public IActionResult ExportCf7(string tn)
        {
            var model = new ExportCf7Model();
            model.SelectedSample = tn;
            model.SetupSampleData(tn);
            return View(model);
        }

        public async Task<IActionResult> GetWorkbookSample5()
        {
            var file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"data\\Allsvenskan2001.xlsx");
            using var package = new ExcelPackage(file);
            var fileBytes = await package.GetAsByteArrayAsync();
            return File(fileBytes, ContentType, "EPPlusHtmlSample5.xlsx");
        }

        public async Task<IActionResult> GetWorkbookSample6()
        {
            var file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"data\\SwedishGeography.xlsx");
            using var package = new ExcelPackage(file);
            var fileBytes = await package.GetAsByteArrayAsync();
            return File(fileBytes, ContentType, "EPPlusHtmlSample6.xlsx");
        }

        public async Task<IActionResult> GetWorkbookSample7()
        {
            var file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"data\\CfExport1.xlsx");
            using var package = new ExcelPackage(file);
            var fileBytes = await package.GetAsByteArrayAsync();
            return File(fileBytes, ContentType, "EPPlusHtmlSample7.xlsx");
        }
    }
}
