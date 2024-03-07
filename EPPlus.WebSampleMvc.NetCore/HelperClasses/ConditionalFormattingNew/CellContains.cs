using EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew
{
    enum BetweenNotBetween
    {
        Between = CellValueCondition.Between,
        Not_Between = CellValueCondition.Not_Between
    }

    enum CellValueConditionsLimited
    {
        Equal_To = CellValueCondition.Equal_To,
        Not_Equal_To = CellValueCondition.Not_Equal_To,
        Greater_Than = CellValueCondition.Greater_Than,
        Less_Than = CellValueCondition.Less_Than,
        Greater_Than_Or_Equal_To = CellValueCondition.Greater_Than_Or_Equal_To,
        Less_Than_Or_Equal_To = CellValueCondition.Less_Than_Or_Equal_To
    }

    internal class CellContains : SimpleFormattingRule
    {
        public override string FormatName => Enum.GetName(CFRuleCategory.CellContains);

        public override CFRuleCategory RuleCategory => CFRuleCategory.CellContains;

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

        private Dictionary<string, DropDown<Enum>> TypesOfCellValue = new Dictionary<string, DropDown<Enum>>();

        private TypeDoubleFormula<BetweenNotBetween> _twoFormulas = new TypeDoubleFormula<BetweenNotBetween>(BetweenNotBetween.Between);
        private TypeFormula<CellValueConditionsLimited> _oneFormula = new TypeFormula<CellValueConditionsLimited>(CellValueConditionsLimited.Equal_To);

        private TypeFormula<SpecificTextCondition> TextDropDown = new TypeFormula<SpecificTextCondition>(SpecificTextCondition.Containing);
        private DropDown<DateCondition> DatesDropDown = new DropDown<DateCondition>(DateCondition.Yesterday);
        public Dictionary<eCellContains, object> DropDownDict = new Dictionary<eCellContains, object>();
        public eCellContains CurrentType = eCellContains.Cell_Value;

        // public Dictionary<eCellContains, DropDownBase> AllOptionsExtra = new Dictionary<eCellContains, DropDownBase>();

        public CellContains()
        {
            string[] emptyArr = new string[0];

            //AllOptionsExtra.Add(eCellContains.Specific_Text, TextDropDown);

            DropDownDict.Add(eCellContains.Cell_Value, FieldValuesDictionary);
            DropDownDict.Add(eCellContains.Specific_Text, TextDropDown);
            DropDownDict.Add(eCellContains.Dates_Occuring, DatesDropDown);
            DropDownDict.Add(eCellContains.Blanks, emptyArr);
            DropDownDict.Add(eCellContains.No_Blanks, emptyArr);
            DropDownDict.Add(eCellContains.Errors, emptyArr);
            DropDownDict.Add(eCellContains.No_Errors, emptyArr);
        }
    }
}
