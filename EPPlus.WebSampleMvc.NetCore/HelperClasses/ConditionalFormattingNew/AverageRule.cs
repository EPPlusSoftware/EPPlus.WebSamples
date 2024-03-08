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
            //Since Averages has to be based on eExcelConditionalFormattingRuleType
            //GetNames makes them appear in a non-sensical order to the user. This rectifies that
            string[] preferredOrder = new string[]
            {
                "Above",
                "Below",
                "Equal_Or_Above",
                "Equal_Or_Below",
                "One_Std_Dev_Above",
                "One_Std_Dev_Below",
                "Two_Std_Dev_Above",
                "Two_Std_Dev_Below",
                "Three_Std_Dev_Above",
                "Three_Std_Dev_Below",
            };

            for(int i = 0; i < preferredOrder.Length; i++) 
            {
                DropDown.Names[i] = preferredOrder[i];
                DropDown.EnumValues[i] = (Averages)Enum.Parse(typeof(Averages), preferredOrder[i]);
            }
        }

    }
}
