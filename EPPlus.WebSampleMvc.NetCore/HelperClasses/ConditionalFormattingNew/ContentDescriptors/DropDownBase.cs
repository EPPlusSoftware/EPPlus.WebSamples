using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors
{
    public abstract class DropDownBase
    {
        public abstract List<string> Names { get; set; }

        public abstract string SelectedName { get; set; }
    }
}
