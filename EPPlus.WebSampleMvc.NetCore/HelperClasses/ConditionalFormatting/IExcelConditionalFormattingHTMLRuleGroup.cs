using EPPlus.WebSampleMvc.NetCore.HelperClasses;
using OfficeOpenXml.ConditionalFormatting;
using System;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormatting
{
    public interface IExcelConditionalFormattingHTMLRuleGroup
    {
        public string FormatTitle { get; }

        public string ListOptionText { get; }

        public string[] Formulas { get; set; }

        public FormatSettings Settings { get; set; }

        public EnumCollection Collection { get; }

        public string ContentLabel { get; }

        public CFRuleType FormatType { get; }

        public bool? CheckBox { get; }

        public ExcelConditionalFormattingRule GetCFForRange(IRangeConditionalFormatting targetRange);
    }
}
