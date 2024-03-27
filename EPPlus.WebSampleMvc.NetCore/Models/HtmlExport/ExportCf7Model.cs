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
                    new() { Name = "3 Color Scale", TableName = "tbl3ColorScale", Description = "A three-color scale conditional formatting of green, yellow and red for values." },
                    new() { Name = "Cell Contains", TableName = "cellContainsTbl", Description = "Example for several 'Format only cells that contain' options. \n" +
                    "Column1: Cell Value Between to show if data is within price range 250-300. \n" +
                    "Column2: Specific Text Containing 'Blue' to show color options \n" +
                    "Column3: Dates Occuring, yesterday, today and tomorrow to highlight update dates in red yellow and green." },
                    new() { Name = "Blanks No blanks", TableName = "blanksTbl", Description = "Example of 'Format only cells that contain' conditional formatting with options blanks and no blanks." },
                    new() { Name = "Errors No Errors", TableName = "errorsTbl" , Description = "Example of 'Format only cells that contain' with options errors and no errors." },
                    new() { Name = "Top 10", TableName = "top10Tbl", Description = "A 'Format only top or bottom ranked values' example with Top10 highlighted in green." },
                    new() { Name = "Above and below average", TableName = "averageTbl", Description = "Two 'Format only values above and below average' conditional formattings. Above average green, below average red." },
                    new() { Name = "Unique and Duplicate", TableName = "uniqueDuplicatesTbl", Description = "A 'Format only unique or duplicate values' example \n " +
                    "Column1 shows duplicates in red.\n" +
                    "Column2 show duplicates in green and uniques in yellow with bold." },
                    new() { Name = "Custom expression", TableName = "expressionTbl", Description = "A 'Use a formula to determine which cells to format' example.\n" +
                    "Column1 with formula 'D45>2002'.\n " +
                    "Column2 with 'isNumber(E45)' and 'isNumber(E45) == FALSE' formulas." },
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
                    Description = sample.Description;
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

                sheet.Calculate();

                Css = exporter.GetCssString();
                Html = exporter.GetHtmlString();
            }
        }

        public string Css { get; set; }

        public string Html { get; set; }

        public string Description { get; set; } = "A three-color scale conditional formatting of green, yellow and red for values.";
    }
}
