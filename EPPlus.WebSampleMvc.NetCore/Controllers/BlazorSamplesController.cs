using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
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
    }
}
