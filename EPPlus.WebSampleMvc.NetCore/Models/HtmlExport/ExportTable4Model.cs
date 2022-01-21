﻿using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using System;
using System.Globalization;
using System.IO;

namespace EPPlus.WebSampleMvc.NetCore.Models.HtmlExport
{
    public class ExportTable4Model
    {
        public void SetupSampleData(TableStyles style = TableStyles.Dark1)
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Html export sample 4");
                var csvFileInfo = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"data\\currencies2011weekly.csv"));
                var format = new ExcelTextFormat
                {
                    Delimiter = ';',
                    Culture = CultureInfo.InvariantCulture,
                    DataTypes = new eDataTypes[] { eDataTypes.DateTime, eDataTypes.Number, eDataTypes.Number, eDataTypes.Number }
                };
                var tableRange = sheet.Cells["A1"].LoadFromText(csvFileInfo, format, style, true);

                sheet.Cells["B1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[tableRange.Start.Row, 1, tableRange.End.Row, 1].Style.Numberformat.Format = "yyyy-MM-dd";
                sheet.Cells[tableRange.Start.Row, 2, tableRange.End.Row, 5].Style.Numberformat.Format = "#,##0.0000";

                var table = sheet.Tables.GetFromRange(tableRange);
                table.ShowFirstColumn = true;

                var settings = table.HtmlExporter.Settings;
                settings.Culture = CultureInfo.InvariantCulture;
                settings.TableId = "currency-table";
                settings.AdditionalTableClassNames.Add("table");
                settings.AdditionalTableClassNames.Add("table-sm");
                settings.AdditionalTableClassNames.Add("table-borderless");

                // export css and html
                Css = table.HtmlExporter.GetCssString();
                Html = table.HtmlExporter.GetHtmlString();
            }
        }

        public string Css { get; set; }

        public string Html { get; set; }
    }
}
