using Microsoft.CodeAnalysis.CSharp.Syntax;
using System;
using System.Collections.Generic;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormatting
{
    abstract public class FormatStyle<T> where T : Enum
    {
        Dictionary<Enum, int> NumFormulas = new Dictionary<Enum, int>();
        List<Enum> DropDowns = new List<Enum>();
    }
}
