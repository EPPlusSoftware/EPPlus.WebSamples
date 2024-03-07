using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors
{
    public class TypeValueColor<T> where T : Enum
    {
        public T CurrentOption { get; set; }
        public string Value { get; set; }
        public Color Color { get; set; }
        public List<string> Names { get; }
        public string SelectedName { get; }

        public TypeValueColor(T type, string value)
        {
            CurrentOption = type;
            Value = value;
            Color = Color;
            Names = GetNames();
            SelectedName = Names[0];
        }

        public List<string> GetNames()
        {
            var names = Enum.GetNames(typeof(T))
                .Select(item => item.Replace('_', ' '));
            return names.ToList();
        }
    }
}
