using System;
using OfficeOpenXml.ConditionalFormatting;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses
{
    [Flags]
    public enum CFRuleCategory
    {
        AllCells = 0,
        CellContains = 2,
        Ranked = 4,
        Average = 8,
        DuplicateUnique = 16,
        CustomExpression = 32,
        All = AllCells | CellContains | Ranked | Average | DuplicateUnique | CustomExpression,
    }

    public enum Ranked
    {
        Top = eExcelConditionalFormattingRuleType.Top,
        Bottom = eExcelConditionalFormattingRuleType.Bottom
    }

    public enum Averages
    {
        Above = eExcelConditionalFormattingRuleType.AboveAverage,
        Below = eExcelConditionalFormattingRuleType.BelowAverage,
        Equal_Or_Above = eExcelConditionalFormattingRuleType.AboveOrEqualAverage,
        Equal_Or_Below = eExcelConditionalFormattingRuleType.BelowOrEqualAverage,
        One_Std_Dev_Above = eExcelConditionalFormattingRuleType.AboveStdDev,
        One_Std_Dev_Below = eExcelConditionalFormattingRuleType.BelowStdDev,
        Two_Std_Dev_Above = eExcelConditionalFormattingRuleType.AboveStdDev,
        Two_Std_Dev_Below = eExcelConditionalFormattingRuleType.BelowStdDev,
        Three_Std_Dev_Above = eExcelConditionalFormattingRuleType.AboveStdDev,
        Three_Std_Dev_Below = eExcelConditionalFormattingRuleType.BelowStdDev,
    }

    public enum DuplicateUnique
    {
        Duplicate = eExcelConditionalFormattingRuleType.DuplicateValues,
        Unique = eExcelConditionalFormattingRuleType.UniqueValues
    }

    public enum CellValueCondition
    {
        Between = eExcelConditionalFormattingRuleType.Between,
        Not_Between = eExcelConditionalFormattingRuleType.NotBetween,
        Equal_To = eExcelConditionalFormattingRuleType.Equal,
        Not_Equal_To = eExcelConditionalFormattingRuleType.NotEqual,
        Greater_Than = eExcelConditionalFormattingRuleType.GreaterThan,
        Less_Than = eExcelConditionalFormattingRuleType.LessThan,
        Greater_Than_Or_Equal_To = eExcelConditionalFormattingRuleType.GreaterThanOrEqual,
        Less_Than_Or_Equal_To = eExcelConditionalFormattingRuleType.LessThanOrEqual
    }

    public enum eCellContains
    {
        Cell_Value,
        Specific_Text,
        Dates_Occuring,
        Blanks,
        No_Blanks,
        Errors,
        No_Errors
    }
    public enum SpecificTextCondition
    {
        Containing = eExcelConditionalFormattingRuleType.ContainsText,
        Not_Containing = eExcelConditionalFormattingRuleType.NotContainsText,
        Beginning_With = eExcelConditionalFormattingRuleType.BeginsWith,
        Ending_With = eExcelConditionalFormattingRuleType.EndsWith
    }
    public enum DateCondition
    {
        Yesterday = eExcelConditionalFormattingRuleType.Yesterday,
        Today = eExcelConditionalFormattingRuleType.Today,
        Tomorrow = eExcelConditionalFormattingRuleType.Tomorrow,
        Last_7_Days = eExcelConditionalFormattingRuleType.Last7Days,
        Last_Week = eExcelConditionalFormattingRuleType.LastWeek,
        This_Week = eExcelConditionalFormattingRuleType.ThisWeek,
        Next_Week = eExcelConditionalFormattingRuleType.NextWeek,
        Last_Month = eExcelConditionalFormattingRuleType.LastMonth,
        This_Month = eExcelConditionalFormattingRuleType.ThisMonth,
        Next_Month = eExcelConditionalFormattingRuleType.NextMonth,
    }

    public enum AdvancedTypes
    {
        Two_Color_Scale = eExcelConditionalFormattingRuleType.TwoColorScale,
        Three_Color_Scale = eExcelConditionalFormattingRuleType.ThreeColorScale,
        Data_Bar = eExcelConditionalFormattingRuleType.DataBar,
        //Icon_Sets
    }

    public enum DataBarTypes
    {
        Automatic,
        Lowest_Value,
        Number,
        Percent,
        Formula,
        Percentile
    }

    public enum TypeScale
    {
        Lowest_Value,
        Number,
        Percent,
        Formula,
        Percentile
    }

    public enum HighVal
    {
        Highest_Value,
        Number,
        Percent,
        Formula,
        Percentile
    }

    public enum TypeScaleMidpoint
    {
        Number = TypeScale.Number,
        Percent = TypeScale.Percent,
        Formula = TypeScale.Formula,
        Percentile = TypeScale.Percentile
    }
}
