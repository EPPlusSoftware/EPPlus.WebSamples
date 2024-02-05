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

                ws.Cells["A1:D11"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells["A1:D11"].Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);

                var cf = GetCFCellContains(ws.Cells["A1:A11"].ConditionalFormatting);

               // var cf = ws.Cells["A1:A11"].ConditionalFormatting.AddBetween();

                cf.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cf.Style.Fill.BackgroundColor.Color = Color.FromKnownColor(AppliedColor);

                cf.Formula = Formula1 == null ? "" : Formula1;
                if (cf.Type == eExcelConditionalFormattingRuleType.Between || cf.Type == eExcelConditionalFormattingRuleType.NotBetween)
                {
                    cf.Formula2 = Formula2 == null ? "" : Formula2;
                    renderFormula2 = true;
                }
                else
                {
                    renderFormula2 = false;
                }

                var exporter = ws.Cells["A1:D11"].CreateHtmlExporter();

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

        public bool renderFormula2 = true;

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
                        //case SpecificTextCondition.Containing:
                        //    break;
                        //case SpecificTextCondition.Not_Containing:
                        //    break;
                        //case SpecificTextCondition.Beginning_With:
                        //    break;
                        //case SpecificTextCondition.Ending_With:
                        //    break;
                        default:
                            throw new NotImplementedException();
                    }
                case CellContains.Dates_Occuring:
                    dateCondition = (DateCondition)DropCF;
                    switch (textCondition)
                    {
                        //case SpecificTextCondition.Containing:
                        //    break;
                        //case SpecificTextCondition.Not_Containing:
                        //    break;
                        //case SpecificTextCondition.Beginning_With:
                        //    break;
                        //case SpecificTextCondition.Ending_With:
                        //    break;
                        default:
                            throw new NotImplementedException();
                    }
                case CellContains.Blanks:
                case CellContains.No_Blanks:
                case CellContains.Errors:
                case CellContains.No_Errors:
                    throw new NotImplementedException();
                default:
                    throw new NotImplementedException();
            }
        }

        public string Css { get; set; }

        public string Html { get; set; }

        //public Array AllBuiltInColors()
        //{
        //    return Enum.GetValues(typeof(Color));
        //}

        //public IEnumerable<SelectListItem> AllBuiltInColors
        //{
        //    get
        //    {
        //        return Enum.GetValues(typeof(KnownColor))
        //            .Cast<KnownColor>()
        //            .Where(x => x == x)
        //            .Select(x => new SelectListItem(x.ToString(), x.ToString()));
        //    }
        //}

        //public IEnumerable<SelectListItem> ValueOperators
        //{
        //    get
        //    {
        //        return Enum.GetValues(typeof(CellValueCondition))
        //            .Cast<CellValueCondition>()
        //            .Where(x => x == x)
        //            .Select(x => new SelectListItem(x.ToString(), x.ToString()));
        //    }
        //}

        //public IEnumerable<SelectListItem> TextConditions
        //{
        //    get
        //    {
        //        return Enum.GetValues(typeof(SpecificTextCondition))
        //            .Cast<SpecificTextCondition>()
        //            .Select(x => new SelectListItem(x.ToString(), x.ToString()));
        //    }
        //}

        //public IEnumerable<SelectListItem> DateConditions
        //{
        //    get
        //    {
        //        return Enum.GetValues(typeof(DateCondition))
        //            .Cast<DateCondition>()
        //            .Select(x => new SelectListItem(x.ToString(), x.ToString()));
        //    }
        //}

        //public IEnumerable<SelectListItem> CellContents
        //{
        //    get
        //    {
        //        return Enum.GetValues(typeof(CellContains))
        //            .Cast<CellContains>()
        //            .Select(x => new SelectListItem(x.ToString(), x.ToString()));
        //    }
        //}

        public string Formula1 { get; set; }
        public string Formula2 { get; set; }

        public CellContains DropOption { get; set; }

        Enum tester;

        public CellValueCondition DropCF { get { return (CellValueCondition)tester; } 
            set 
            { 
                tester = value; 
            } 
        }

        public string DropCFSting { get { return JsonSerializer.Serialize(DropCF.ToString()); } } 

        public KnownColor AppliedColor { get; set; }

        private CellContains _cellContains = CellContains.Cell_Value;

        public Array GetEnumValues()
        {
            return Enum.GetValues(typeof(CellContains));
        }

        public Array GetColValues()
        {
            return Enum.GetValues(typeof(KnownColor));
        }

        //public CellContains cellContains 
        //{   get { return _cellContains; } 
        //    set 
        //    {
        //        _cellContains = value;
        //        enumVariable = GetEnumVariable(); 
        //    } 
        //}

        public CellValueCondition valueCondition { get; set; }

        public DateCondition dateCondition { get; set; }

        public SpecificTextCondition textCondition { get; set; }


        //public Enum GetEnumVariable()
        //{
        //    switch (cellContains)
        //    {
        //        case CellContains.Cell_Value:
        //            return valueCondition;
        //        case CellContains.Specific_Text:
        //            return textCondition;
        //        case CellContains.Dates_Occuring:
        //            return dateCondition;
        //        case CellContains.Blanks:
        //        case CellContains.No_Blanks:
        //        case CellContains.Errors:
        //        case CellContains.No_Errors:
        //            return null;
        //        default: return null;
        //    }
        //}

        //public Enum GetChildEnum()
        //{
        //    switch (DropCF)
        //    {
        //        case CellContains.Cell_Value:
        //            return (CellValueCondition)DropCF;
        //        case CellContains.Specific_Text:
        //            return (SpecificTextCondition)DropCF;
        //        case CellContains.Dates_Occuring:
        //            return (DateCondition)DropCF;
        //        case CellContains.Blanks:
        //        case CellContains.No_Blanks:
        //        case CellContains.Errors:
        //        case CellContains.No_Errors:
        //            return null;
        //    }
        //    return null;
        //}

        //internal IEnumerable<SelectListItem> GetConditionList()
        //{
        //    switch (cellContains)
        //    {
        //        case CellContains.Cell_Value:
        //            return ValueOperators;
        //        case CellContains.Specific_Text:
        //            return TextConditions;
        //        case CellContains.Dates_Occuring:
        //            return DateConditions;
        //        case CellContains.Blanks:
        //        case CellContains.No_Blanks:
        //        case CellContains.Errors:
        //        case CellContains.No_Errors:
        //            return null;
        //        default: return null;
        //    }
        //}
    }
}
