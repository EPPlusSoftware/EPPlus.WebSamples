using OfficeOpenXml.ConditionalFormatting;
using System;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormatting.Formats
{
    public abstract class FormatBase : IExcelConditionalFormattingHTMLRuleGroup
    {
        abstract public string ListOptionText { get; }
        abstract public string FormatTitle { get; }

        abstract public CFRuleType FormatType { get; }

        abstract public EnumCollection Collection { get; protected set; }

        abstract public string[] Formulas { get; set; }

        abstract public FormatSettings Settings { get; set; }

        abstract public string ContentLabel { get; }

        abstract public bool? CheckBox { get; protected set; }
        abstract public ExcelConditionalFormattingRule GetCFForRange(IRangeConditionalFormatting targetRange);
    }
}
