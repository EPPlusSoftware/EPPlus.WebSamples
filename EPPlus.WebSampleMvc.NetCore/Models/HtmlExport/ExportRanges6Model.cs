using OfficeOpenXml;
using System;
using System.IO;

namespace EPPlus.WebSampleMvc.NetCore.Models.HtmlExport
{
    public class ExportRanges6Model
    {
        public void SetupSampleData()
        {
            using(var package = new ExcelPackage(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"data\\SwedishGeography.xlsx")))
            {
                var sheet = package.Workbook.Worksheets[0];
                var exporter = package.Workbook.CreateHtmlExporter(sheet.Cells["A1:D10"], sheet.Cells["G1:J10"]);
                exporter.Settings.Minify = false;
                Html1 = exporter.GetHtmlString(0, x =>
                {
                    x.TableId = "cities-table";
                    x.Accessibility.TableSettings.AriaLabel = "Largest cities in Sweden";
                });
                Html2 = exporter.GetHtmlString(1, x =>
                {
                    x.TableId = "lakes-table";
                    x.Accessibility.TableSettings.AriaLabel = "Largest lakes in Sweden";
                });
                Css = exporter.GetCssString();
            }
        }

        public string Css { get; set; }

        public string Html1 { get; set; }

        public string Html2 { get; set; }
    }
}
