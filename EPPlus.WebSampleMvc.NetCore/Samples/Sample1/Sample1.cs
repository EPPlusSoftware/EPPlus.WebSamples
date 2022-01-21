using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.Samples.Sample1
{
    public static class Sample1
    {
        public static void InitWorkbook(ExcelPackage package)
        {
            var sheet = package.Workbook.Worksheets.Add("Sheet 1");
            sheet.Cells["A1:C1"].Merge = true;
            sheet.Cells["A1"].Style.Font.Size = 18f;
            sheet.Cells["A1"].Style.Font.Bold = true;
            sheet.Cells["A1"].Value = "Simple example 1";
        }
    }
}
