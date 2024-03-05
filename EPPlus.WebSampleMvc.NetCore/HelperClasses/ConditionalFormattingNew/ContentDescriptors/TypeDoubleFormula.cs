using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors
{
    public class TypeDoubleFormula<T> : TypeFormula<T> where T : Enum
    {
        public string Formula2 { get; set; } = "";

        public TypeDoubleFormula(T type) : base(type)
        {
        }
    }
}
