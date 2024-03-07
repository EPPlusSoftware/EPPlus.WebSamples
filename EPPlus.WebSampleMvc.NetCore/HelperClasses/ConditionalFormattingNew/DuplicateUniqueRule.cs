using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew
{
    public class DuplicateUniqueRule : SimpleFormattingRule
    {
        public override string FormatName => Enum.GetName(CFRuleCategory.UniqueDuplicates);
        public override CFRuleCategory RuleCategory => CFRuleCategory.UniqueDuplicates;

        public override string ListOptionText => "Format only unique or duplicate values";

        public override string FormatTitle => "Format all";

        public override string ContentLabel => "values in the selected range";

        public DropDown<DuplicateUnique> DropDown = new DropDown<DuplicateUnique>(DuplicateUnique.Duplicate);

        public DuplicateUniqueRule()
        {

        }
    }
}
