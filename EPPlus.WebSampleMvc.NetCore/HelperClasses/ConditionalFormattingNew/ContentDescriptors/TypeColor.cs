using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors
{
    public class TypeColor<T> : DropDown<T> where T : Enum
    {
        public Color ColorProp { get; set; }

        public TypeColor(T type) : base(type)
        {
            ColorProp = Color.Blue;
        }
        public TypeColor(T type, Color color) : base(type)
        {
            ColorProp = color;
        }
    }
}
