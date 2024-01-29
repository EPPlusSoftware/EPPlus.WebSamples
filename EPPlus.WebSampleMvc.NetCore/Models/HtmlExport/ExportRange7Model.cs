using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport;
using System;
using System.Drawing;
using System.Linq;
using System.IO;
using Microsoft.AspNetCore.Mvc.Rendering;
using System.Collections.Generic;


namespace EPPlus.WebSampleMvc.NetCore.Models.HtmlExport
{
    public class ExportRange7Model
    {
        public void SetupSampleData(KnownColor col = KnownColor.Goldenrod, string Formula1 = "-3", string Formula2 = "3")
        {
            using (var package = new ExcelPackage(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"data\\SwedishGeography.xlsx")))
            {
                var ws = package.Workbook.Worksheets.Add("CF_ws");

                ws.Cells["A1:A11"].Formula = "ROW()-5";

                ws.Cells["A1:D11"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells["A1:D11"].Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);


                var cfBetween = ws.Cells["A1:A11"].ConditionalFormatting.AddBetween();

                cfBetween.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cfBetween.Style.Fill.BackgroundColor.Color = Color.FromKnownColor(col);

                cfBetween.Formula = Formula1;
                cfBetween.Formula2 = Formula2;

                var exporter = ws.Cells["A1:D11"].CreateHtmlExporter();

                var settings = exporter.Settings;
                settings.Pictures.Include = ePictureInclude.Include;
                //settings.Pictures.KeepOriginalSize = true;
                settings.Minify = false;
                settings.SetColumnWidth = true;
                settings.SetRowHeight = true;
                settings.Pictures.AddNameAsId = true;

                ws.Calculate();

                // export css and html
                Css = exporter.GetCssString();
                Html = exporter.GetHtmlString();
            }
        }

        public string Css { get; set; }

        public string Html { get; set; }

        public IEnumerable<SelectListItem> AllBuiltInColors
        {
            get
            {
                return Enum.GetValues(typeof(KnownColor))
                    .Cast<KnownColor>()
                    .Where(x => x == x)
                    .Select(x => new SelectListItem(x.ToString(), x.ToString()));
            }
        }

        public IEnumerable<SelectListItem> ValueOperators
        {
            get
            {
                return Enum.GetValues(typeof(CellValueOperator))
                    .Cast<CellValueOperator>()
                    .Where(x => x == x)
                    .Select(x => new SelectListItem(x.ToString(), x.ToString()));
            }
        }

        public IEnumerable<SelectListItem> WithTypes
        {
            get
            {
                return new SelectListItem[] { new SelectListItem { Text = "Cell Value" } };
            }
        }

        public string Formula1 { get; set; }
        public string Formula2 { get; set; }

        public KnownColor bgColBetween { get; set; }

        public CellValueOperator valueOperator { get; set; }

        public string withType =  "Cell Value";

        public enum CellValueOperator
        {
            Between,
            Not_Between,
            Equal_To,
            Not_Equal_To,
            Greater_Than,
            Less_Than,
            Greater_Than_Or_Equal_To,
            Less_Than_Or_Equal_To
        }
    }
}
