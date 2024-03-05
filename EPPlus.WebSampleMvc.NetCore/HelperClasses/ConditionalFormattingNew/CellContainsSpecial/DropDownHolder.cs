using EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.CellContainsSpecial
{
    public class DropDownHolder<T, U> : DropDown<T> where T : Enum where U : Enum
    {
        public TypeFormula<U> TypeFormula { get; set; }

        public DropDownHolder(T type, U type2) : base(type)
        {
            TypeFormula = new TypeFormula<U>(type2);
        }
    }
}
