using OfficeOpenXml.ConditionalFormatting;
using System;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormatting.Formats
{
    public class UniqueDuplicateFormat : FormatBase
    {
        public override string ListOptionText => "Format only unique or duplicate values";

        public override string FormatTitle => "Format all";

        public override CFRuleType FormatType => CFRuleType.UniqueDuplicates;

        public override EnumCollection Collection { get; protected set; }
        public override string[] Formulas { get; set; }
        public override FormatSettings Settings { get; set; }

        public override string ContentLabel => "values in the selected range";

        public override bool? CheckBox { get; protected set; }

        public UniqueDuplicateFormat() : this(new InputData())
        {
        }

        public UniqueDuplicateFormat(InputData input)
        {
            Collection = new EnumCollection(typeof(UniqueDuplicate), new Type[0], input.SelectedEnums);
            Settings = input.Settings;
        }

        public override ExcelConditionalFormattingRule GetCFForRange(IRangeConditionalFormatting targetRange)
        {
            if(Collection.SelectedKey == "Duplicate")
            {
                return (ExcelConditionalFormattingRule)targetRange.AddDuplicateValues();
            }
            else
            {
                return (ExcelConditionalFormattingRule)targetRange.AddUniqueValues();
            }
        }
    }
}
