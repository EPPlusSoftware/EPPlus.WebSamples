using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport;
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
                var sheet2 = package.Workbook.Worksheets[1];
                var table = package.Workbook.Worksheets[2].Tables["Municipalities"];
                var exporter = package.Workbook.CreateHtmlExporter(sheet.Cells["A1:D10"], sheet.Cells["G1:J10"], sheet2.Cells["A1:E73"], table.Range);
                exporter.Settings.Pictures.Include = ePictureInclude.Include;
                exporter.Settings.Minify = false;
                // all hyperlinks should open in a new browser window.
                exporter.Settings.HyperlinkTarget = "_blank";
                // change name of the data-value attribute to data-sort
                // to make the dataset sortable in DataTables.js.
                exporter.Settings.DataValueAttributeName = "sort";
                Css = exporter.GetCssStringAsync().Result;
                Html1 = exporter.GetHtmlStringAsync(0, x =>
                {
                    x.TableId = "cities-table";
                    x.Accessibility.TableSettings.AriaLabel = "Largest cities in Sweden";
                }).Result;
                Html3 = exporter.GetHtmlStringAsync(2, x =>
                {
                    x.TableId = "kings-table";
                    x.Accessibility.TableSettings.AriaLabel = "Some kings in Sweden";
                }).Result;
                Html2 = exporter.GetHtmlStringAsync(1, x =>
                {
                    x.TableId = "lakes-table";
                    x.Accessibility.TableSettings.AriaLabel = "Largest lakes in Sweden";
                }).Result;
                Html4 = exporter.GetHtmlStringAsync(3, x =>
                {
                    x.TableId = "municipalities-table";
                    x.Accessibility.TableSettings.AriaLabel = "Swedish municipalities";
                    // Add bootstrap table classes
                    x.AdditionalTableClassNames.Add("table");
                    x.AdditionalTableClassNames.Add("table-sm");
                }).Result;
            }
        }

        public string Css { get; set; }

        public string Html1 { get; set; }

        public string Html2 { get; set; }

        public string Html3 { get; set; }

        public string Html4 { get; set; }
    }
}
