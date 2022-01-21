using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.Samples.Sample2
{
    public static class Sample2
    {
        private static OrderRepository _orderRepository = new OrderRepository();
        
        /// <summary>
        /// Creates a worksheet and loads 1000 instances of <see cref="Order"></see> into that worksheet.
        /// </summary>
        /// <param name="package">The <see cref="ExcelPackage"/></param>
        public static void InitWorkbook(ExcelPackage package)
        {
            var orders = _orderRepository.GetAllOrders();
            var sheet = package.Workbook.Worksheets.Add("Orders");
            
            // Create the header row
            sheet.Cells["A1:C1"].Merge = true;
            sheet.Cells["A1"].Style.Font.Bold = true;
            sheet.Cells["A1"].Style.Font.Size = 18f;
            sheet.Cells["A1"].Value = "Orders";
            
            // load Order objects into the worksheet
            sheet.Cells["A3"].LoadFromCollection(orders, PrintHeaders: true, TableStyle: OfficeOpenXml.Table.TableStyles.Light12);
            // autofit columns to get the correct width
            sheet.Cells[3, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column].AutoFitColumns();
            // set style
            sheet.Cells[3, 1, sheet.Dimension.End.Row, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
            sheet.Cells[4, 3, sheet.Dimension.End.Row, 3].Style.Numberformat.Format = "yyyy-mm-dd";
        }
    }
}
