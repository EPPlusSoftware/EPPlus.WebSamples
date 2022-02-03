using System.Data;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace EPPlus.WebSampleMvc.NetCore.Models.HtmlExport
{
    public class ExportTable2Model
    {
        public ExportTable2Model()
        {

        }

        private DataTable _dataTable;

        private void InitDataTable()
        {
            _dataTable = new DataTable();
            _dataTable.Columns.Add("Country", typeof(string));
            _dataTable.Columns.Add("Population", typeof(int));
            var areaCol = _dataTable.Columns.Add("Area", typeof(int));
            areaCol.Caption = "Area (km2)";


            _dataTable.Rows.Add("Sweden", 10409248, 450295);
            _dataTable.Rows.Add("Norway", 5402171, 385178);
            _dataTable.Rows.Add("Netherlands", 17553530, 41198);
            _dataTable.Rows.Add("Finland", 5541806, 338145);
            _dataTable.Rows.Add("Belgium", 11521238, 30510);
            _dataTable.Rows.Add("Denmark", 5850189, 44493);
            _dataTable.Rows.Add("Lithuania", 2801264, 65300);
            _dataTable.Rows.Add("Greece", 10718565, 131940);
            _dataTable.Rows.Add("Russia", 145734038, 3972400);
            _dataTable.Rows.Add("Germany", 83124418, 357386);
            _dataTable.Rows.Add("France", 64990511, 551695);
            _dataTable.Rows.Add("Czech Republic", 10665677, 78866);
            _dataTable.Rows.Add("Slovakia", 5459781, 49036);
            _dataTable.Rows.Add("Spain", 47394223, 498468);
            _dataTable.Rows.Add("Portugal", 10256193, 91568);
            _dataTable.Rows.Add("United Kingdom", 67141684, 242495);
            _dataTable.Rows.Add("Poland", 37921592, 312685);
            _dataTable.Rows.Add("Albania", 2882740, 28748);
            _dataTable.Rows.Add("Estonia", 1322920, 45339);
            _dataTable.Rows.Add("Hungary", 9707499, 93030);
            _dataTable.Rows.Add("Romania", 19186000, 238397);
            _dataTable.Rows.Add("Italy", 60627291, 301338);
            _dataTable.Rows.Add("Bulgaria", 7051608, 110994);
            _dataTable.Rows.Add("Belarus", 9452617, 207600);
            _dataTable.Rows.Add("Austria", 8891388, 83858);
            _dataTable.Rows.Add("Switzerland", 8525611, 41290);
            _dataTable.Rows.Add("Ireland", 4818690, 70273);
            _dataTable.Rows.Add("Ukraine", 44246156, 603628);
            _dataTable.Rows.Add("Iceland", 336713, 102775);
            _dataTable.Rows.Add("Serbia", 6871547, 77453);
            _dataTable.Rows.Add("Croatia", 4156405, 56594);
            _dataTable.Rows.Add("Latvia", 1928459, 64589);
            _dataTable.Rows.Add("Bosnia and Herzegovina", 3323925, 51129);
            _dataTable.Rows.Add("Montenegro", 627809, 13812);
            _dataTable.Rows.Add("Cyrprus", 1189265, 9251);
            _dataTable.Rows.Add("Kosovo", 1798506, 10908);
            _dataTable.Rows.Add("Slovenia", 2055496, 20273);
            _dataTable.Rows.Add("Moldova", 4033963, 33846);
            _dataTable.Rows.Add("North Macedonia", 2083374, 25713);

        }

        public void SetupSampleData(TableStyles style = TableStyles.Dark1)
        {
            InitDataTable();
            using(var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Html export sample 2");
                var tableRange = sheet.Cells["A1"].LoadFromDataTable(_dataTable, true, style);
                
                // configure the table
                var table = sheet.Tables.GetFromRange(tableRange);
                table.Sort(x => x.SortBy.ColumnNamed("Population", eSortOrder.Descending));
                table.ShowTotal = true;
                table.Columns[0].TotalsRowLabel = "Total";
                table.Columns[1].TotalsRowFunction = RowFunctions.Sum;
                table.Columns[2].TotalsRowFunction = RowFunctions.Sum;
                
                // add column for population density
                table.Columns.Add(1);
                tableRange = table.Range;
                table.Columns[3].CalculatedColumnFormula = $"{table.Name}[[#This Row],[Population]]/{table.Name}[[#This Row],[Area (km2)]]";
                table.Columns[3].Name = "Density";
                table.Columns[3].TotalsRowFunction = RowFunctions.Average;
                sheet.Calculate();

                // format the header
                sheet.Cells[1, tableRange.Start.Column + 1, 1, tableRange.End.Column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                // format the rows
                var lastDataRow = tableRange.End.Row - 1;
                sheet.Cells[tableRange.Start.Row, 1, lastDataRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells[tableRange.Start.Row, 2, lastDataRow, 2].Style.Numberformat.Format = "#,##0";
                sheet.Cells[tableRange.Start.Row, 3, lastDataRow, 3].Style.Numberformat.Format = "#,##0 \"km2\"";
                sheet.Cells[tableRange.Start.Row, 4, lastDataRow, 4].Style.Numberformat.Format = "#,##0.0";

                // format the total row
                var totalRow = tableRange.End.Row;
                sheet.Cells[totalRow, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells[totalRow, 2].Style.Numberformat.Format = "#,##0";
                sheet.Cells[totalRow, 3].Style.Numberformat.Format = "#,##0 \"km2\"";
                sheet.Cells[totalRow, 4].Style.Numberformat.Format = "\"Avg: \"#,##0.0 ";


                // Configure export settings
                var exporter = table.CreateHtmlExporter();
                var settings = exporter.Settings;
                settings.Minify = false;
                settings.AdditionalTableClassNames.Add("table");
                settings.AdditionalTableClassNames.Add("table-sm");
                settings.TableId = "population-table";
                settings.Accessibility.TableSettings.AriaDescribedBy = "table-description";
                settings.Accessibility.TableSettings.AriaLabel = "Demo table";

                // export css and html
                Css = exporter.GetCssString();
                Html = exporter.GetHtmlString();
            }
        }

        public string Css { get; set; }

        public string Html { get; set; }


    }
}
