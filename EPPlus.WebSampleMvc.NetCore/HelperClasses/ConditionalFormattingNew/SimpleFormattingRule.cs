using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew
{
    public abstract class SimpleFormattingRule : FormattingRuleBase
    {
        public abstract string ContentLabel { get; }
        //public abstract FormatSettings Settings { get; set; }
    }
}
