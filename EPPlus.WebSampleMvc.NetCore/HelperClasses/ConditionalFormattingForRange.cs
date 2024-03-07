using OfficeOpenXml.ConditionalFormatting;
using System;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using System.Drawing;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses
{
    public static class ConditionalFormattingForRange
    {
        public static ExcelConditionalFormattingRule GetCFForRangeClass(IRangeConditionalFormatting targetRange, eExcelConditionalFormattingRuleType type)
        {
            return (ExcelConditionalFormattingRule)GetCFForRange(targetRange, type);
        }

        public static IExcelConditionalFormattingRule GetCFForRange(IRangeConditionalFormatting targetRange, eExcelConditionalFormattingRuleType type)
        {
            switch (type)
            {
                case eExcelConditionalFormattingRuleType.GreaterThan:
                    return targetRange.AddGreaterThan();
                case eExcelConditionalFormattingRuleType.GreaterThanOrEqual:
                    return targetRange.AddGreaterThanOrEqual();
                case eExcelConditionalFormattingRuleType.LessThan:
                    return targetRange.AddLessThan();
                case eExcelConditionalFormattingRuleType.LessThanOrEqual:
                    return targetRange.AddLessThanOrEqual();
                case eExcelConditionalFormattingRuleType.AboveAverage:
                    return targetRange.AddAboveAverage();
                case eExcelConditionalFormattingRuleType.AboveOrEqualAverage:
                    return targetRange.AddAboveOrEqualAverage();
                case eExcelConditionalFormattingRuleType.BelowAverage:
                    return targetRange.AddBelowAverage();
                case eExcelConditionalFormattingRuleType.BelowOrEqualAverage:
                    return targetRange.AddBelowOrEqualAverage();
                case eExcelConditionalFormattingRuleType.AboveStdDev:
                    return targetRange.AddAboveStdDev();
                case eExcelConditionalFormattingRuleType.BelowStdDev:
                    return targetRange.AddBelowStdDev();
                case eExcelConditionalFormattingRuleType.Bottom:
                    return targetRange.AddBottom();
                case eExcelConditionalFormattingRuleType.BottomPercent:
                    return targetRange.AddBottomPercent();
                case eExcelConditionalFormattingRuleType.Top:
                    return targetRange.AddTop();
                case eExcelConditionalFormattingRuleType.TopPercent:
                    return targetRange.AddTopPercent();
                case eExcelConditionalFormattingRuleType.Last7Days:
                    return targetRange.AddLast7Days();
                case eExcelConditionalFormattingRuleType.LastMonth:
                    return targetRange.AddLastMonth();
                case eExcelConditionalFormattingRuleType.LastWeek:
                    return targetRange.AddLastWeek();
                case eExcelConditionalFormattingRuleType.NextMonth:
                    return targetRange.AddNextMonth();
                case eExcelConditionalFormattingRuleType.NextWeek:
                    return targetRange.AddNextWeek();
                case eExcelConditionalFormattingRuleType.ThisMonth:
                    return targetRange.AddThisMonth();
                case eExcelConditionalFormattingRuleType.ThisWeek:
                    return targetRange.AddThisWeek();
                case eExcelConditionalFormattingRuleType.Today:
                    return targetRange.AddToday();
                case eExcelConditionalFormattingRuleType.Tomorrow:
                    return targetRange.AddTomorrow();
                case eExcelConditionalFormattingRuleType.Yesterday:
                    return targetRange.AddTomorrow();
                case eExcelConditionalFormattingRuleType.BeginsWith:
                    return targetRange.AddBeginsWith();
                case eExcelConditionalFormattingRuleType.Between:
                    return targetRange.AddBetween();
                case eExcelConditionalFormattingRuleType.ContainsBlanks:
                    return targetRange.AddContainsBlanks();
                case eExcelConditionalFormattingRuleType.ContainsErrors:
                    return targetRange.AddContainsErrors();
                case eExcelConditionalFormattingRuleType.ContainsText:
                    return targetRange.AddContainsText();
                case eExcelConditionalFormattingRuleType.DuplicateValues:
                    return targetRange.AddDuplicateValues();
                case eExcelConditionalFormattingRuleType.EndsWith:
                    return targetRange.AddEndsWith();
                case eExcelConditionalFormattingRuleType.Equal:
                    return targetRange.AddEqual();
                case eExcelConditionalFormattingRuleType.Expression:
                    return targetRange.AddExpression();
                case eExcelConditionalFormattingRuleType.NotBetween:
                    return targetRange.AddNotBetween();
                case eExcelConditionalFormattingRuleType.NotContainsBlanks:
                    return targetRange.AddNotContainsBlanks();
                case eExcelConditionalFormattingRuleType.NotContainsErrors:
                    return targetRange.AddNotContainsErrors();
                case eExcelConditionalFormattingRuleType.NotContainsText:
                    return targetRange.AddNotContainsText();
                case eExcelConditionalFormattingRuleType.NotEqual:
                    return targetRange.AddNotEqual();
                case eExcelConditionalFormattingRuleType.UniqueValues:
                    return targetRange.AddUniqueValues();
                case eExcelConditionalFormattingRuleType.ThreeColorScale:
                    return targetRange.AddThreeColorScale();
                case eExcelConditionalFormattingRuleType.TwoColorScale:
                    return targetRange.AddTwoColorScale();
                case eExcelConditionalFormattingRuleType.ThreeIconSet:
                    return targetRange.AddThreeIconSet(eExcelconditionalFormatting3IconsSetType.Arrows);
                case eExcelConditionalFormattingRuleType.FourIconSet:
                    return targetRange.AddFourIconSet(eExcelconditionalFormatting4IconsSetType.Arrows);
                case eExcelConditionalFormattingRuleType.FiveIconSet:
                    return targetRange.AddFiveIconSet(eExcelconditionalFormatting5IconsSetType.Arrows);
                case eExcelConditionalFormattingRuleType.DataBar:
                    return targetRange.AddDatabar(Color.Blue);
                default:
                    throw new NotImplementedException($"Not implemented/impossible type: {type} was picked");
            }
        }
    }
}
