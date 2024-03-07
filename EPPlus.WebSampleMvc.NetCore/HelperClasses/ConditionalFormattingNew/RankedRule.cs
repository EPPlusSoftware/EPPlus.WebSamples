using EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew
{
    internal class RankedRule : SimpleFormattingRule
    {
        public override string FormatName => Enum.GetName(CFRuleCategory.Ranked);

        public override string ListOptionText => "Format only top or bottom ranked values";
        public override string FormatTitle => "Format values that rank in the:";
        public override string ContentLabel => "% of the selected range";

        public bool PercentOfRange = false;

        public DropDown<Ranked> DropDown { get; set; } = new DropDown<Ranked>(Ranked.Top);

        public override CFRuleCategory RuleCategory => CFRuleCategory.Ranked;

        internal RankedRule()
        {
        }

    }
}
