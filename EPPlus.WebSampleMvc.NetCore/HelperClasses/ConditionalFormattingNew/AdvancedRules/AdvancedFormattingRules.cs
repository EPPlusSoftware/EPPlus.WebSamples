using System;
using System.Collections.Generic;
using EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.AdvancedRules
{
    public class AdvancedFormattingRules : FormattingRuleBase
    {
        public override string ListOptionText => "Format all cells based on their values";

        public override string FormatTitle => ListOptionText;

        public DropDown<AdvancedTypes> DropDown = new DropDown<AdvancedTypes>(AdvancedTypes.Two_Color_Scale);

        public string FormatName => Enum.GetName(CFRuleType.AllCells);

        //internal override string ContentLabel => throw new NotImplementedException();

        public Dictionary<AdvancedTypes, object> Types = new Dictionary<AdvancedTypes, object>
        {
            {AdvancedTypes.Two_Color_Scale, new TwoColorScale() },
            {AdvancedTypes.Three_Color_Scale, new ThreeColorScale() },
            {AdvancedTypes.Data_Bar, new DataBar() },
        };

        public AdvancedFormattingRules() { }

        public AdvancedFormattingRules(InputData input)
        {
            
        }

    }
}
