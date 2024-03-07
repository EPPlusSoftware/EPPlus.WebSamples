using OfficeOpenXml.ConditionalFormatting;
using System;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormatting.Formats
{
    public class AverageFormat : FormatBase
    {
        public override string ListOptionText => "Format only values that are above or below average";

        public override string FormatTitle => "Format values that are";

        public override CFRuleCategory FormatType => CFRuleCategory.Average;

        public override EnumCollection Collection { get; protected set; }
        public override string[] Formulas { get; set; }
        public override FormatSettings Settings { get; set; }

        public override string ContentLabel => "the average for the selected range";

        public override bool? CheckBox { get; protected set; }

        public AverageFormat() : this(new InputData())
        {
        }

        public AverageFormat(InputData input)
        {
            Collection = new EnumCollection(typeof(Averages), new Type[0], input.SelectedEnums);
            Settings = input.Settings;
        }

        public override ExcelConditionalFormattingRule GetCFForRange(IRangeConditionalFormatting targetRange)
        {
            switch (Collection.SelectedKey)
            {
                case "Above":
                    return (ExcelConditionalFormattingRule)targetRange.AddAboveAverage();
                case "Below":
                    return (ExcelConditionalFormattingRule)targetRange.AddBelowAverage();
                case "Equal_Or_Above":
                    return (ExcelConditionalFormattingRule)targetRange.AddAboveOrEqualAverage();
                case "Equal_Or_Below":
                    return (ExcelConditionalFormattingRule)targetRange.AddBelowOrEqualAverage();
                case "One_Std_Dev_Above":
                case "Two_Std_Dev_Above":
                case "Three_Std_Dev_Above":
                    var cf = targetRange.AddAboveStdDev();
                    if (Collection.SelectedKey[2] == 'e')
                    {
                        cf.StdDev = 1;
                    }
                    else if(Collection.SelectedKey[2] == 'o')
                    {
                        cf.StdDev = 2;
                    }
                    else
                    {
                        cf.StdDev = 3;
                    }
                    return (ExcelConditionalFormattingRule)cf;
                case "One_Std_Dev_Below":
                case "Two_Std_Dev_Below":
                case "Three_Std_Dev_Below":
                    var cf2 = targetRange.AddBelowStdDev();
                    if (Collection.SelectedKey[2] == 'e')
                    {
                        cf2.StdDev = 1;
                    }
                    else if (Collection.SelectedKey[2] == 'o')
                    {
                        cf2.StdDev = 2;
                    }
                    else
                    {
                        cf2.StdDev = 3;
                    }
                    return (ExcelConditionalFormattingRule)cf2;
                default: throw new InvalidOperationException($"{Collection.SelectedKey} is not a valid option");
            }
        }
    }
}
