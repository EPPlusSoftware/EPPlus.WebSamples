using OfficeOpenXml.ConditionalFormatting;
using System;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormatting.Formats
{
    public class ExpressionFormat : FormatBase
    {
        public override string ListOptionText => "Use a Formula to determine which cells to format";

        public override string FormatTitle => "Format values where this formula is true";

        public override CFRuleCategory FormatType => CFRuleCategory.CustomExpression;

        public override EnumCollection Collection { get { return null; } protected set { throw new NotImplementedException("should never be set"); } }
        public override string[] Formulas { get; set; }
        public override FormatSettings Settings { get; set; }

        public override string ContentLabel => "";

        public override bool? CheckBox { get; protected set; }

        public ExpressionFormat() : this(new InputData())
        {
        }

        public ExpressionFormat(InputData input) 
        {
            Formulas = input.Formulas;
            Settings = input.Settings;
        }

        public override ExcelConditionalFormattingRule GetCFForRange(IRangeConditionalFormatting targetRange)
        {
            return (ExcelConditionalFormattingRule)targetRange.AddExpression();
        }
    }
}
