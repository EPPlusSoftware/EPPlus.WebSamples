using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.Controllers
{
    public class BlazorSamplesController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Sample1()
        {
            return View();
        }

        [HttpGet, Route("/api/fxrates")]
        public IActionResult FxRates()
        {
            var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"data\\FxRates.json");
            var json = System.IO.File.ReadAllText(path);
            return Json(json);
        }
    }
}
