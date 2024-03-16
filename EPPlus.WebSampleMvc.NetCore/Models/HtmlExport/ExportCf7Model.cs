using OfficeOpenXml;
using System.IO;
using System;
using System.Globalization;
using System.Collections;
using EPPlus.WebSampleMvc.NetCore.Models.HtmlExport.ConditionalFormattings;
using System.Collections.Generic;

namespace EPPlus.WebSampleMvc.NetCore.Models.HtmlExport
{
    public class ExportCf7Model
    {
        public ExportCf7Model()
        {
            
        }

        public const string DefaultSample = "tbl3ColorScale";

        public string SelectedSample { get; set; }

        public IEnumerable<CfSample> AllSamples
        {
            get
            {
                return new List<CfSample>
                {
                    new() { Name = "3 Color Scale", TableName = "tbl3ColorScale" },
                    new() { Name = "Data Bars", TableName = "tblDataBars" },
                    new() { Name = "Expression", TableName = "tblExpression" },

                };
            }
        }

        private string GetTableName(string tableName)
        {
            if(string.IsNullOrWhiteSpace(tableName))
            {
                return DefaultSample;
            }
            foreach(var sample in AllSamples)
            {
                if(string.Compare(sample.TableName, tableName, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    return sample.TableName;
                }
            }
            return DefaultSample;
        }

        public void SetupSampleData(string tableName = "")
        {
            var tblName = GetTableName(tableName);
            using (var package = new ExcelPackage(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"data\\CfExport1.xlsx")))
            {
                var sheet = package.Workbook.Worksheets[0];
                var table = sheet.Tables[tblName];
                var exporter = package.Workbook.CreateHtmlExporter(table.Range);
                var settings = exporter.Settings;
                settings.Minify = false;
                settings.Culture = CultureInfo.InvariantCulture;
                settings.TableId = "cf-table";
                settings.SetColumnWidth = true;
                settings.SetRowHeight = true;


                Css = exporter.GetCssString();
                Html = exporter.GetHtmlString();
            }
        }

        public string Css { get; set; }

        public string Html { get; set; }
    }
}
