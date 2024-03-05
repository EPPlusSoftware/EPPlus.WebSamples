using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors
{
    public class TypeFormula<T> : DropDown<T> where T : Enum
    {
        public string Formula { get; set; }

        public TypeFormula(T type) : base(type)
        {
            Formula = "";
        }
    }
}
