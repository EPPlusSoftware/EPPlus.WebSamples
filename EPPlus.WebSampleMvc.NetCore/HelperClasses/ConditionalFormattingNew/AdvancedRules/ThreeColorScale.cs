using EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.AdvancedRules
{
    internal class ThreeColorScale : TwoColorScale
    {
        internal TypeValueColor<TypeScaleMidpoint> Midpoint { get; set; }

        internal ThreeColorScale()
        {
            Midpoint = new TypeValueColor<TypeScaleMidpoint>(TypeScaleMidpoint.Percentile, "");
        }
    }
}
