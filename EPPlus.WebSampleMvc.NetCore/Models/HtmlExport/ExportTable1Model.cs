using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Faker;
using Microsoft.AspNetCore.Mvc.Rendering;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EPPlus.WebSampleMvc.NetCore.Models.HtmlExport
{
    public class ExportTable1Model
    {
        public ExportTable1Model()
        {

        }

        private DataTable _dataTable;

        private void InitDataTable()
        {
            _dataTable = new DataTable();
            _dataTable.Columns.Add("Country", typeof(string));
            _dataTable.Columns.Add("FirstName", typeof(string));
            _dataTable.Columns.Add("LastName", typeof(string));
            _dataTable.Columns.Add("BirthDate", typeof(DateTime));
            _dataTable.Columns.Add("City", typeof(string));
            

            for(var x = 0; x < 50; x++)
            {
                _dataTable.Rows.Add(Faker.Address.UkCountry(), Faker.Name.First(), Faker.Name.Last(), Faker.Identification.DateOfBirth(), Faker.Address.City());
            }

        }

        public void SetupSampleData(int theme, TableStyles? style = TableStyles.Dark1)
        {
            // This method just fakes some data into a data table
            InitDataTable();

            using(var package = new ExcelPackage())
            {
                SetTheme(theme, package);

                var sheet = package.Workbook.Worksheets.Add("Html export sample 1");
                var tableRange = sheet.Cells["A1"].LoadFromDataTable(_dataTable, true, style);
                // set number format for the BirthDate column
                sheet.Cells[tableRange.Start.Row + 1, 4, tableRange.End.Row, 4].Style.Numberformat.Format = "yyyy-MM-dd";
                tableRange.AutoFitColumns();

                var table = sheet.Tables.GetFromRange(tableRange);

                // table properties
                table.ShowFirstColumn = this.ShowFirstColumn;
                table.ShowLastColumn = this.ShowLastColumn;
                table.ShowColumnStripes = this.ShowColumnStripes;
                table.ShowRowStripes = this.ShowRowsStripes;

                // Export Html and CSS
                var exporter = table.CreateHtmlExporter();
                exporter.Settings.Minify = false;
                Css = exporter.GetCssString();
                Html = exporter.GetHtmlString();
                WorkbookBytes = package.GetAsByteArray();
            }
        }

        private static void SetTheme(int theme, ExcelPackage package)
        {
            if (theme > 0)
            {
                var fileInfo = default(FileInfo);
                switch (theme)
                {
                    case 1:
                        fileInfo = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"themes\\Ion.thmx"));
                        break;
                    case 2:
                        fileInfo = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"themes\\Banded.thmx"));
                        break;
                    case 3:
                        fileInfo = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"themes\\Parallax.thmx"));
                        break;
                    default:
                        fileInfo = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"themes\\Ion.thmx"));
                        break;
                }
                package.Workbook.ThemeManager.Load(fileInfo);
            }
        }

        public IEnumerable<SelectListItem> AllBuiltInTableStyles
        {
            get
            {
                return System.Enum.GetValues(typeof(TableStyles))
                    .Cast<TableStyles>()
                    .Where(x => x != TableStyles.Custom)
                    .Select(x => new SelectListItem(x.ToString(), x.ToString()));
            }
        }

        public IEnumerable<SelectListItem> AllThemes
        {
            get
            {
                return new List<SelectListItem>
                {
                    new SelectListItem("Default (Office)", "0"),
                    new SelectListItem("Ion", "1"),
                    new SelectListItem("Banded", "2"),
                    new SelectListItem("Parallax", "3")
                };
            }
        }

        public bool ShowFirstColumn { get; set; }

        public bool ShowLastColumn { get; set; }

        public bool ShowColumnStripes { get; set; }

        public bool ShowRowsStripes { get; set; }

        public string TableStyle { get; set; }

        public int Theme { get; set; }

        public bool AddBootstrapClasses { get; set; }

        public bool AddDataTablesJs { get; set; }

        public bool GetWorkbook { get; set; }

        public string Css { get; set; }

        public string Html { get; set; }

        public byte[] WorkbookBytes { get; set; }


    }
}
