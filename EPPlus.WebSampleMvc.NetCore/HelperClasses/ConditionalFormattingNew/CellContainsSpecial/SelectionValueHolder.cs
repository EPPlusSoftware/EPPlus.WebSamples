using Microsoft.Extensions.Primitives;
using System;
using System.Drawing;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.CellContainsSpecial
{
    public class SelectionValueHolder
    {
        public Enum Type { get; set; }

        public string[] Values { get; set; }

        public SelectionValueHolder(Enum type, string[] values) 
        {
            Type = type;
            Values = values;
        }
    }
}
