using EPPlus.WebSampleMvc.NetCore.HelperClasses;
using EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormatting;
using EPPlus.WebSampleMvc.NetCore.Models.HtmlExport;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Drawing;
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
        public IActionResult ExportRange7()
        {
            var model = new ExportRange7Model
            {
                CurrentRuleType = CFRuleType.CellContains,
                CurrentRuleTypeStr = "CellContains",
                Formula1 = "3",
                Formula2 = "-3",
                Settings = new FormatSettings(CFColor.Blue),
            };
            model.SetupSampleData();
            return View(model);
        }

        [HttpPost, ValidateAntiForgeryToken]
        public IActionResult ExportRange7(ExportRange7Model model)
        {
            model.SetupSampleData();
            return View(model);
        }

        //[HttpPost, Route("/api/types/{model, formula, index}")]
        //public void UpdateFormula(ExportRange7Model model, string formula, int index)
        //{
        //    if (index == 1)
        //    {
        //        model.Formula1 = formula;
        //    }
        //    else if (index == 2)
        //    {
        //        model.Formula2 = formula;
        //    }
        //    else
        //    {
        //        throw new NotImplementedException();
        //    }
        //}

        //[HttpPost, Route("/api/types/update")]
        //public JsonResult UpdateFormula1(string dropA, string dropB)
        //{

        //    //model.Formula1 = changedFormula;
        //    //return Json(Ok(model.Formula1));
        //}

        //[HttpPost, Route("/api/types/{model}/{test}")]
        //public JsonResult UpdateFormula1(ExportRange7Model model, string changedFormula)
        //{
        //    model.Formula1 = changedFormula;
        //    return Json(Ok(model.Formula1));
        //}

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
    }
}
