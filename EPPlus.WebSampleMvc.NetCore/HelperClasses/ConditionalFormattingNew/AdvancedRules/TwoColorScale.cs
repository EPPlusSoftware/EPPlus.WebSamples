using EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.AdvancedRules
{
    public class TwoColorScale
    {
        public TypeValueColor<TypeScale> Minimum { get; set; }
        public TypeValueColor<HighVal> Maximum { get; set; }

        public TwoColorScale()
        {
            Minimum = new TypeValueColor<TypeScale>(TypeScale.Lowest_Value, "");
            Maximum = new TypeValueColor<HighVal>(HighVal.Highest_Value, "");
        }
    }
}
