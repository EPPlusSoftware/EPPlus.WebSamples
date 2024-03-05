using EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew
{
    internal class CellContains : SimpleFormattingRule
    {
        public string FormatName => Enum.GetName(CFRuleType.CellContains);

        public override string ListOptionText { get => "Format only cells that contain"; }
        public override string FormatTitle { get => "Format only cells with:"; }
        public override string ContentLabel { get => "and"; }

        private Dictionary<CellValueCondition, string[]> FieldValuesDictionary =
                new Dictionary<CellValueCondition, string[]>()
        {
                    {CellValueCondition.Between, new string[2] },
                    {CellValueCondition.Not_Between, new string[2] },
                    {CellValueCondition.Equal_To, new string[1] },
                    {CellValueCondition.Not_Equal_To, new string[1] },
                    {CellValueCondition.Greater_Than, new string[1] },
                    {CellValueCondition.Less_Than, new string[1] },
                    {CellValueCondition.Greater_Than_Or_Equal_To, new string[1] },
                    {CellValueCondition.Less_Than_Or_Equal_To, new string[1] },
        };

        private TypeFormula<SpecificTextCondition> TextDropDown = new TypeFormula<SpecificTextCondition>(SpecificTextCondition.Containing);
        private DropDown<DateCondition> DatesDropDown = new DropDown<DateCondition>(DateCondition.Yesterday);
        public Dictionary<eCellContains, object> AllOptions = new Dictionary<eCellContains, object>();

        // public Dictionary<eCellContains, DropDownBase> AllOptionsExtra = new Dictionary<eCellContains, DropDownBase>();

        public CellContains()
        {
            string[] emptyArr = new string[0];

            //AllOptionsExtra.Add(eCellContains.Specific_Text, TextDropDown);

            AllOptions.Add(eCellContains.Cell_Value, FieldValuesDictionary);
            AllOptions.Add(eCellContains.Specific_Text, TextDropDown);
            AllOptions.Add(eCellContains.Dates_Occuring, DatesDropDown);
            AllOptions.Add(eCellContains.Blanks, emptyArr);
            AllOptions.Add(eCellContains.No_Blanks, emptyArr);
            AllOptions.Add(eCellContains.Errors, emptyArr);
            AllOptions.Add(eCellContains.No_Errors, emptyArr);
        }
    }
}
