using EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors;
using System;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.CellContainsSpecial
{
    public class TwoFormula<T> where T : Enum
    {
        internal TypeFormula<T> NormalValues;
        internal TypeFormula<T> SpecialCaseValues;

        public TwoFormula(T defaultType)
        {
            NormalValues = new TypeFormula<T>(defaultType);
            NormalValues.Formula = "1";
            SpecialCaseValues = new TypeFormula<T>(defaultType);
            SpecialCaseValues.Formula = "2";
        }
    }
}
