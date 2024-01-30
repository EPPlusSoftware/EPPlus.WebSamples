﻿using OfficeOpenXml;
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
        public void SetupSampleData(KnownColor col = KnownColor.Goldenrod, string Formula1 = "-3", string Formula2 = "3")
        {
            using (var package = new ExcelPackage(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"data\\SwedishGeography.xlsx")))
            {
                var ws = package.Workbook.Worksheets.Add("CF_ws");

                ws.Cells["A1:A11"].Formula = "ROW()-5";

                ws.Cells["A1:D11"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                ws.Cells["A1:D11"].Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);

                var cf = GetCFCellContains(ws.Cells["A1:A11"].ConditionalFormatting);

                cf.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                cf.Style.Fill.BackgroundColor.Color = Color.FromKnownColor(col);

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

        ExcelConditionalFormattingRule GetCFCellContains(IRangeConditionalFormatting targetRange)
        {
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
        }

        public string Css { get; set; }

        public string Html { get; set; }

        public IEnumerable<SelectListItem> AllBuiltInColors
        {
            get
            {
                return Enum.GetValues(typeof(KnownColor))
                    .Cast<KnownColor>()
                    .Where(x => x == x)
                    .Select(x => new SelectListItem(x.ToString(), x.ToString()));
            }
        }

        public IEnumerable<SelectListItem> ValueOperators
        {
            get
            {
                return Enum.GetValues(typeof(CellValueCondition))
                    .Cast<CellValueCondition>()
                    .Where(x => x == x)
                    .Select(x => new SelectListItem(x.ToString(), x.ToString()));
            }
        }

        public IEnumerable<SelectListItem> TextConditions
        {
            get
            {
                return Enum.GetValues(typeof(SpecificTextCondition))
                    .Cast<SpecificTextCondition>()
                    .Select(x => new SelectListItem(x.ToString(), x.ToString()));
            }
        }

        public IEnumerable<SelectListItem> DateConditions
        {
            get
            {
                return Enum.GetValues(typeof(DateCondition))
                    .Cast<DateCondition>()
                    .Select(x => new SelectListItem(x.ToString(), x.ToString()));
            }
        }

        public IEnumerable<SelectListItem> CellContents
        {
            get
            {
                return Enum.GetValues(typeof(CellContains))
                    .Cast<CellContains>()
                    .Select(x => new SelectListItem(x.ToString(), x.ToString()));
            }
        }

        public string Formula1 { get; set; }
        public string Formula2 { get; set; }

        public KnownColor bgColBetween { get; set; }

        private CellContains _cellContains = CellContains.Cell_Value;

        public CellContains cellContains 
        {   get { return _cellContains; } 
            set 
            {
                _cellContains = value;
                enumVariable = GetEnumVariable(); 
            } 
        }

        public CellValueCondition valueCondition { get; set; }

        public DateCondition dateCondition { get; set; }

        public SpecificTextCondition textCondition { get; set; }

        public Enum enumVariable { get; set; }

        public Enum GetEnumVariable()
        {
            switch (cellContains)
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
                    return null;
                default: return null;
            }
        }

        internal IEnumerable<SelectListItem> GetConditionList()
        {

            switch (cellContains)
            {
                case CellContains.Cell_Value:
                    return ValueOperators;
                case CellContains.Specific_Text:
                    return TextConditions;
                case CellContains.Dates_Occuring:
                    return DateConditions;
                case CellContains.Blanks:
                case CellContains.No_Blanks:
                case CellContains.Errors:
                case CellContains.No_Errors:
                    return null;
                default: return null;
            }
        }
    }
}
