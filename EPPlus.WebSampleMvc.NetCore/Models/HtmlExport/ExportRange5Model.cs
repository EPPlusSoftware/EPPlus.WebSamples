using OfficeOpenXml;
using System;
using System.Globalization;
using System.IO;

namespace EPPlus.WebSampleMvc.NetCore.Models.HtmlExport
{
    public class ExportRange5Model
    {
        public void SetupSampleData()
        {
            var file = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"data\\Allsvenskan2001.xlsx");
            using (var package = new ExcelPackage(file))
            {
                var sheet = package.Workbook.Worksheets[0];

                var exporter = sheet.Cells["B5:N19"].CreateHtmlExporter();
                var settings = exporter.Settings;
                settings.Culture = CultureInfo.InvariantCulture;
                settings.TableId = "soccer-table";
                settings.Accessibility.TableSettings.AriaLabel = "This html-table is a range that is exported from EPPlus";

                // use column width from the workbook
                settings.SetColumnWidth = true;
                settings.SetRowHeight = true;
                
                // when Minify is false the output will be formated with indentation and linebreaks.
                settings.Minify = false;

                // export css and html
                Css = exporter.GetCssString();
                Html = exporter.GetHtmlString();
            }
        }

        public string Css { get; set; }

        public string Html { get; set; }
    }
}
