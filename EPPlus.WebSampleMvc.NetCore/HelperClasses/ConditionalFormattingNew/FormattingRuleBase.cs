using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew
{
    public abstract class FormattingRuleBase
    {
        public abstract string ListOptionText { get; }
        public abstract string FormatTitle { get; }
        public abstract string FormatName { get; }
        public abstract CFRuleCategory RuleCategory { get; }
    }
}
