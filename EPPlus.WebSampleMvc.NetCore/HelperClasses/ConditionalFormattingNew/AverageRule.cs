using EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew
{
    public class AverageRule : SimpleFormattingRule
    {
        public override string ListOptionText => "Format only values that are above or below average";
        public override string FormatTitle => "Format values that are";
        public override string ContentLabel => "the average for the selected range";

        public override string FormatName => Enum.GetName(CFRuleCategory.Average);
        public override CFRuleCategory RuleCategory => CFRuleCategory.Average;

        public DropDown<Averages> DropDown { get; set; } = new DropDown<Averages>(Averages.Above);

        public AverageRule()
        {

        }

    }
}
