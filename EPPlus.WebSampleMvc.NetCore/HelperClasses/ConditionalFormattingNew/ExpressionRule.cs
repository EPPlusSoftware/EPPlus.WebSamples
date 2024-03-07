using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew
{
    public class ExpressionRule : SimpleFormattingRule
    {
        public override string FormatName => Enum.GetName(CFRuleCategory.CustomExpression);
        public override CFRuleCategory RuleCategory => CFRuleCategory.CustomExpression;

        public override string ListOptionText => "Use a Formula to determine which cells to format";

        public override string FormatTitle => "Format values where this formula is true";

        public override string ContentLabel => "";

        //public override FormatSettings Settings { get; set; }

        public string Formula = "";

        public ExpressionRule()
        {

        }
    }
}
