using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport;
using System;
using System.Drawing;
using System.Linq;
using System.IO;
using Microsoft.AspNetCore.Mvc.Rendering;
using System.Collections.Generic;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils.Extensions;
using EPPlus.WebSampleMvc.NetCore.HelperClasses;
using System.Text.Json;
using OfficeOpenXml.Style;
using System.Globalization;



namespace EPPlus.WebSampleMvc.NetCore.Models.HtmlExport
{
    public enum CellContains
    {
        Cell_Value,
        Specific_Text,
        Dates_Occuring,
        Blanks,
        No_Blanks,
        Errors,
        No_Errors
    }

    public enum CellValueCondition
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

    public enum SpecificTextCondition
    {
        Containing,
        Not_Containing,
        Beginning_With,
        Ending_With
    }

    public enum DateCondition
    {
        Yesterday,
        Today,
        Tomorrow,
        Last_7_Days,
        Last_Week,
        This_Week,
        Next_Week,
        Last_Month,
        This_Month,
        Next_Month
    }


    public class ExportRange7Model
    {
        public void SetupSampleData()
        {
            using (var package = new ExcelPackage())
            {
                var collection = new EnumCollection<CellContains>(new Type[] {typeof(CellValueCondition), typeof(SpecificTextCondition), typeof(DateCondition)});
                JsonData = JsonSerializer.Serialize(collection.dict);

                var ws = package.Workbook.Worksheets.Add("CF_ws");

                ws.Cells["A1:A11"].Formula = "ROW()-5";
                ws.Cells["B1:B11"].Formula = "\"Row number \" & ROW()";
                ws.Cells["C1:C11"].Formula = "TODAY() -2 + ROW()";
                ws.Cells["D1"].Formula = "TODAY() -6";
                ws.Cells["D2"].Formula = "TODAY()";
                ws.Cells["D3"].Formula = "TODAY() + 7";
                ws.Cells["D5"].Formula = "TODAY() -31";
                ws.Cells["D6"].Formula = "TODAY()";
                ws.Cells["D7"].Formula = "TODAY() + 31";
                ws.Cells["E1:E5"].Formula = "E10-E11+NULL";


                ws.Cells["C1:D11"].Style.Numberformat.Format = "mm-dd-yy";

                ws.Cells["A1:E11"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells["A1:E11"].Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);

                var cf = GetCFCellContains(ws.Cells["A1:E11"].ConditionalFormatting);

               // var cf = ws.Cells["A1:A11"].ConditionalFormatting.AddBetween();

                cf.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cf.Style.Fill.BackgroundColor.Color = Color.FromKnownColor(AppliedColor);

                cf.Formula = Formula1 == null ? "" : Formula1;
                if (cf.Type == eExcelConditionalFormattingRuleType.Between || cf.Type == eExcelConditionalFormattingRuleType.NotBetween)
                {
                    cf.Formula2 = Formula2 == null ? "" : Formula2;
                }

                var exporter = ws.Cells["A1:E11"].CreateHtmlExporter();

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

        public string JsonData = "";

        ExcelConditionalFormattingRule GetCFCellContains(IRangeConditionalFormatting targetRange)
        {
            switch (DropOption)
            {
                case CellContains.Cell_Value:
                    valueCondition = (CellValueCondition)DropCF;
                    switch (valueCondition)
                    {
                        case CellValueCondition.Between:
                            return (ExcelConditionalFormattingRule)targetRange.AddBetween();
                        case CellValueCondition.Not_Between:
                            return (ExcelConditionalFormattingRule)targetRange.AddNotBetween();
                        case CellValueCondition.Equal_To:
                            return (ExcelConditionalFormattingRule)targetRange.AddEqual();
                        case CellValueCondition.Not_Equal_To:
                            return (ExcelConditionalFormattingRule)targetRange.AddNotEqual();
                        case CellValueCondition.Greater_Than:
                            return (ExcelConditionalFormattingRule)targetRange.AddGreaterThan();
                        case CellValueCondition.Less_Than:
                            return (ExcelConditionalFormattingRule)targetRange.AddLessThan();
                        case CellValueCondition.Greater_Than_Or_Equal_To:
                            return (ExcelConditionalFormattingRule)targetRange.AddGreaterThanOrEqual();
                        case CellValueCondition.Less_Than_Or_Equal_To:
                            return (ExcelConditionalFormattingRule)targetRange.AddLessThanOrEqual();
                        default:
                            throw new NotImplementedException();
                    }
                case CellContains.Specific_Text:
                    textCondition = (SpecificTextCondition)DropCF;
                    switch (textCondition)
                    {
                        case SpecificTextCondition.Containing:
                            return (ExcelConditionalFormattingRule)targetRange.AddContainsText();
                        case SpecificTextCondition.Not_Containing:
                            return (ExcelConditionalFormattingRule)targetRange.AddNotContainsText();
                        case SpecificTextCondition.Beginning_With:
                            return (ExcelConditionalFormattingRule)targetRange.AddBeginsWith();
                        case SpecificTextCondition.Ending_With:
                            return (ExcelConditionalFormattingRule)targetRange.AddEndsWith();
                        default:
                            throw new NotImplementedException();
                    }
                case CellContains.Dates_Occuring:
                    dateCondition = (DateCondition)DropCF;
                    switch (dateCondition)
                    {
                        case DateCondition.Yesterday:
                            return (ExcelConditionalFormattingRule)targetRange.AddYesterday();
                        case DateCondition.Today:
                            return (ExcelConditionalFormattingRule)targetRange.AddToday();
                        case DateCondition.Tomorrow:
                            return (ExcelConditionalFormattingRule)targetRange.AddTomorrow();
                        case DateCondition.Last_7_Days:
                            return (ExcelConditionalFormattingRule)targetRange.AddLast7Days();
                        case DateCondition.Last_Week:
                            return (ExcelConditionalFormattingRule)targetRange.AddLastWeek();
                        case DateCondition.This_Week:
                            return (ExcelConditionalFormattingRule)targetRange.AddThisWeek();
                        case DateCondition.Next_Week:
                            return (ExcelConditionalFormattingRule)targetRange.AddNextWeek();
                        case DateCondition.Last_Month:
                            return (ExcelConditionalFormattingRule)targetRange.AddLastMonth();
                        case DateCondition.This_Month:
                            return (ExcelConditionalFormattingRule)targetRange.AddThisMonth();
                        case DateCondition.Next_Month:
                            return (ExcelConditionalFormattingRule)targetRange.AddNextMonth();
                        default:
                            throw new NotImplementedException();
                    }
                case CellContains.Blanks:
                    return (ExcelConditionalFormattingRule)targetRange.AddContainsBlanks();
                case CellContains.No_Blanks:
                    return (ExcelConditionalFormattingRule)targetRange.AddNotContainsBlanks();
                case CellContains.Errors:
                    return (ExcelConditionalFormattingRule)targetRange.AddContainsErrors();
                case CellContains.No_Errors:
                    return (ExcelConditionalFormattingRule)targetRange.AddNotContainsErrors();
                default:
                    throw new NotImplementedException();
            }
        }

        public string Css { get; set; }

        public string Html { get; set; }

        public string Formula1 { get; set; }
        public string Formula2 { get; set; }

        public CellContains DropOption { get; set; }

        public Enum DropCF { get { return GetCFEnumVariable(); } set { SetCFVariable(value); } }

        public string DropCFSting { get { return GetSerializedDropB(); } } 

        public KnownColor AppliedColor { get; set; }

        public Array GetEnumValues()
        {
            return Enum.GetValues(typeof(CellContains));
        }

        public Array GetColValues()
        {
            return Enum.GetValues(typeof(KnownColor));
        }

        public CellValueCondition valueCondition { get; set; }

        public DateCondition dateCondition { get; set; }

        public SpecificTextCondition textCondition { get; set; }

        private Enum GetCFEnumVariable()
        {
            switch (DropOption)
            {
                case CellContains.Cell_Value:
                    return valueCondition;
                case CellContains.Specific_Text:
                    return textCondition;
                case CellContains.Dates_Occuring:
                    return dateCondition;
                case CellContains.Blanks:
                case CellContains.No_Blanks:
                case CellContains.Errors:
                case CellContains.No_Errors:
                default:
                    return null;
            }
        }

        private string GetSerializedDropB()
        {
            switch (DropOption)
            {
                case CellContains.Cell_Value:
                    return JsonSerializer.Serialize(Enum.GetName(valueCondition));
                case CellContains.Specific_Text:
                    return JsonSerializer.Serialize(Enum.GetName(textCondition));
                case CellContains.Dates_Occuring:
                    return JsonSerializer.Serialize(Enum.GetName(dateCondition));
                case CellContains.Blanks:
                case CellContains.No_Blanks:
                case CellContains.Errors:
                case CellContains.No_Errors:
                default:
                    return JsonSerializer.Serialize("");
            }
        }

        private void SetCFVariable(object aValue)
        {
            switch (DropOption)
            {
                case CellContains.Cell_Value:
                    valueCondition = (CellValueCondition)aValue;
                    break;
                case CellContains.Specific_Text:
                    textCondition = (SpecificTextCondition)aValue;
                    break;
                case CellContains.Dates_Occuring:
                    dateCondition = (DateCondition)aValue;
                    break;
                case CellContains.Blanks:
                case CellContains.No_Blanks:
                case CellContains.Errors:
                case CellContains.No_Errors:
                    break;
            }
        }
    }
}
