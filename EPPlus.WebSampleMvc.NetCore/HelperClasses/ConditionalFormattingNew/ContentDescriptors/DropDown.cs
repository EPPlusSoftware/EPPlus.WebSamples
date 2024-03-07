using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors
{
    public class DropDown<T> : DropDownBase where T : Enum
    {
        public T Selected { get; set; }

        public override List<string> Names { get; }

        public override string SelectedName { get; set; } = "";

        public List<T> EnumValues { get; }

        public DropDown(T type)
        {
            Selected = type;
            Names = GetNames();
            SelectedName = Names[0];
            EnumValues = GetEnumValues();
        }

        public List<string> GetNames()
        {
            //var names = Enum.GetNames(typeof(T))
            //    .Select(item => item.Replace('_', ' '));
            return Enum.GetNames(typeof(T)).ToList();
        }

        private List<T> GetEnumValues() 
        {
            return Enum.GetValues(typeof(T)).Cast<T>().ToList();
        }
    }
}
