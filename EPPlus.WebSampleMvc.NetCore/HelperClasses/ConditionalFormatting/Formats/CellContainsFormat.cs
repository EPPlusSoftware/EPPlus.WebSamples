using System.Drawing;
using System.Text.Json;
using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormatting.Formats
{
    public class CellContainsFormat : FormatBase
    {
        public override string ListOptionText { get => "Format only cells that contain"; }
        public override string FormatTitle { get => "Format only cells with:"; }
        public override string ContentLabel { get => "and"; }
        public override EnumCollection Collection { get; protected set; }
        public override string[] Formulas { get; set; }
        public override FormatSettings Settings { get; set; }
        public override CFRuleCategory FormatType => CFRuleCategory.CellContains;
        public override bool? CheckBox { get; protected set; } = null;

        public CellContainsFormat() : this(new InputData())
        {
        }

        public CellContainsFormat(InputData input)
        {
            Collection = new EnumCollection(typeof(eCellContains), new Type[] { typeof(CellValueCondition), typeof(SpecificTextCondition), typeof(DateCondition) }, input.SelectedEnums);

            Settings = input.Settings;
            if (input.Formulas.Length < 2)
            {
                Formulas = new string[] { input.Formulas[0], "" };
            }
            else
            {
                Formulas = input.Formulas;
            }
        }

        public override ExcelConditionalFormattingRule GetCFForRange(IRangeConditionalFormatting targetRange)
        {
            switch (Collection.SelectedKey)
            {
                case "Cell_Value":
                    switch (Collection.SelectedValue)
                    {
                        case "Between":
                            return (ExcelConditionalFormattingRule)targetRange.AddBetween();
                        case "Not_Between":
                            return (ExcelConditionalFormattingRule)targetRange.AddNotBetween();
                        case "Equal_To":
                            return (ExcelConditionalFormattingRule)targetRange.AddEqual();
                        case "Not_Equal_To":
                            return (ExcelConditionalFormattingRule)targetRange.AddNotEqual();
                        case "Greater_Than":
                            return (ExcelConditionalFormattingRule)targetRange.AddGreaterThan();
                        case "Less_Than":
                            return (ExcelConditionalFormattingRule)targetRange.AddLessThan();
                        case "Greater_Than_Or_Equal_To":
                            return (ExcelConditionalFormattingRule)targetRange.AddGreaterThanOrEqual();
                        case "Less_Than_Or_Equal_To":
                            return (ExcelConditionalFormattingRule)targetRange.AddLessThanOrEqual();
                        default:
                            throw new NotImplementedException();
                    }
                case "Specific_Text":
                    switch (Collection.SelectedValue)
                    {
                        case "Containing":
                            return (ExcelConditionalFormattingRule)targetRange.AddContainsText();
                        case "Not_Containing":
                            return (ExcelConditionalFormattingRule)targetRange.AddNotContainsText();
                        case "Beginning_With":
                            return (ExcelConditionalFormattingRule)targetRange.AddBeginsWith();
                        case "Ending_With":
                            return (ExcelConditionalFormattingRule)targetRange.AddEndsWith();
                        default:
                            throw new NotImplementedException();
                    }
                case "Dates_Occuring":
                    switch (Collection.SelectedValue)
                    {
                        case "Yesterday":
                            return (ExcelConditionalFormattingRule)targetRange.AddYesterday();
                        case "Today":
                            return (ExcelConditionalFormattingRule)targetRange.AddToday();
                        case "Tomorrow":
                            return (ExcelConditionalFormattingRule)targetRange.AddTomorrow();
                        case "Last_7_Days":
                            return (ExcelConditionalFormattingRule)targetRange.AddLast7Days();
                        case "Last_Week":
                            return (ExcelConditionalFormattingRule)targetRange.AddLastWeek();
                        case "This_Week":
                            return (ExcelConditionalFormattingRule)targetRange.AddThisWeek();
                        case "Next_Week":
                            return (ExcelConditionalFormattingRule)targetRange.AddNextWeek();
                        case "Last_Month":
                            return (ExcelConditionalFormattingRule)targetRange.AddLastMonth();
                        case "This_Month":
                            return (ExcelConditionalFormattingRule)targetRange.AddThisMonth();
                        case "Next_Month":
                            return (ExcelConditionalFormattingRule)targetRange.AddNextMonth();
                        default:
                            throw new NotImplementedException();
                    }
                case "Blanks":
                    return (ExcelConditionalFormattingRule)targetRange.AddContainsBlanks();
                case "No_Blanks":
                    return (ExcelConditionalFormattingRule)targetRange.AddNotContainsBlanks();
                case "Errors":
                    return (ExcelConditionalFormattingRule)targetRange.AddContainsErrors();
                case "No_Errors":
                    return (ExcelConditionalFormattingRule)targetRange.AddNotContainsErrors();
                default:
                    throw new NotImplementedException();
            }
        }

        public Dictionary<string, int> lookup = new Dictionary<string, int>()
        {
            {"Between", 2},
            {"NotBetween", 2},
            {"Equal_To", 1},
            {"Not_Equal_To", 1}
        };

        public override int ActiveFormulaFields(string selectedValue)
        {
            switch(selectedValue)
            {
                case "Between":
                case "Not_Between":
                    return 2;
                case "Equal_To":
                case "Not_Equal_To":
                case "Greater_Than":
                case "Less_Than":
                case "Greater_Than_Or_Equal_To":
                case "Less_Than_Or_Equal_To":
                case "Containing":
                case "Not_Containing":
                case "Beginning_With":
                case "Ending_With":
                    return 1;
                case "Yesterday":
                case "Today":
                case "Tomorrow":
                case "Last_7_Days":
                case "Last_Week":
                case "This_Week":
                case "Next_Week":
                case "Last_Month":
                case "This_Month":
                case "Next_Month":
                case "Blanks":
                case "No_Blanks":
                case "Errors":
                case "No_Errors":
                    return 0;
                default: return 0;
            }
        }


        //public CellContainsFormat(CFColor color, string[] formulas, FormatSettings settings)
        //{
        //    Initalize();
        //    Settings.SetColor(color);
        //}

        //public override bool Equals(object? obj)
        //{
        //    return base.Equals(obj);
        //}

        //public override int GetHashCode()
        //{
        //    return base.GetHashCode();
        //}

        //public override string? ToString()
        //{
        //    return base.ToString();
        //}
    }
}
