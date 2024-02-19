using System.Drawing;
using System.Text.Json;
using OfficeOpenXml.ConditionalFormatting;
using System;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormatting.Formats
{
    public class RankedFormat : FormatBase
    {
        public override string ListOptionText => "Format only top or bottom ranked values";

        public override string FormatTitle => "Format values that rank in the:";

        public override string ContentLabel => "% of the selected range";

        public override EnumCollection Collection { get; protected set; }

        public override CFRuleType FormatType => CFRuleType.Ranked;

        public override string[] Formulas { get; set; }
        public override FormatSettings Settings { get; set; }

        public override bool? CheckBox { get; protected set; } = false;

        public RankedFormat() : this(new InputData())
        {
        }

        public RankedFormat(InputData input)
        {
            Collection = new EnumCollection(typeof(Ranked), new System.Type[0], input.SelectedEnums);
            Settings = input.Settings;
            Formulas = new string[1] { input.Formulas[0] };

            CheckBox = input.CheckBox;

            //if (input.Formulas == null)
            //{
            //    Formulas = [""];
            //    _asPercent = false;
            //}
            //else
            //{
            //    var formulas = input.Formulas;
            //    Formulas = [formulas[0] == null ? "" : formulas[0]];
            //    _asPercent = formulas[1] == "1" ? true : false;
            //}
        }

        public override ExcelConditionalFormattingRule GetCFForRange(IRangeConditionalFormatting targetRange)
        {
            if(CheckBox == null)
            {
                CheckBox = false;
            }

            switch (Collection.SelectedKey)
            {
                case "Top":
                    if(CheckBox.Value)
                    {
                        return (ExcelConditionalFormattingRule)targetRange.AddTopPercent();
                    }
                    return (ExcelConditionalFormattingRule)targetRange.AddTop();
                case "Bottom":
                    if (CheckBox.Value)
                    {
                        return (ExcelConditionalFormattingRule)targetRange.AddBottomPercent();
                    }
                    return (ExcelConditionalFormattingRule)targetRange.AddBottom();

                default: throw new InvalidOperationException();
            }
        }
    }
}
