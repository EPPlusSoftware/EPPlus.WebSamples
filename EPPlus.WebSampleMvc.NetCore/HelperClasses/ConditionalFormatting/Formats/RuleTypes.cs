using System;
using System.Collections.Generic;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormatting.Formats
{
    public class RuleTypes
    {
        //Dictionary<CFRuleType, IExcelConditionalFormattingHTMLRuleGroup> Types = new Dictionary<CFRuleType, IExcelConditionalFormattingHTMLRuleGroup>();
        public List<IExcelConditionalFormattingHTMLRuleGroup> Types;

        public IExcelConditionalFormattingHTMLRuleGroup ActiveRule;
        //Dictionary<CFRuleType, IExcelConditionalFormattingHTMLRuleGroup?> dict = new Dictionary<CFRuleType, IExcelConditionalFormattingHTMLRuleGroup?>()
        //{
        //    { CFRuleType.CellContains, null},
        //};

        public RuleTypes()
        {
            Types = new List<IExcelConditionalFormattingHTMLRuleGroup>();
            foreach (var item in Enum.GetValues(typeof(CFRuleType)))
            {
                InitializeTypes((CFRuleType)item);
            }
            ActiveRule = Types[0];
        }

        public RuleTypes(CFRuleType ruleType, string[] selectedEnums = null, string[] formulas = null, FormatSettings settings = null, bool? checkbox = null)
        {
            Types = new List<IExcelConditionalFormattingHTMLRuleGroup>();
            var input = new InputData(selectedEnums, formulas, settings, checkbox);

            switch (ruleType)
            {
                case CFRuleType.CellContains:
                    Types.Add(new CellContainsFormat(input));
                    break;
                case CFRuleType.AllCells:
                    break;
                case CFRuleType.Ranked:
                    Types.Add(new RankedFormat(input));
                    break;
                case CFRuleType.Average:
                    break;
                case CFRuleType.UniqueDuplicates:
                    break;
                case CFRuleType.CustomExpression:
                    break;
            }

            ActiveRule = Types[0];

            foreach (var item in Enum.GetValues(typeof(CFRuleType)))
            {
                if ((CFRuleType)item != ruleType)
                {
                    InitializeTypes((CFRuleType)item);
                }
            }
        }

        void InitializeTypes(CFRuleType ruleType)
        {
            switch (ruleType)
            {
                case CFRuleType.CellContains:
                    Types.Add(new CellContainsFormat());
                    break;
                case CFRuleType.AllCells:
                    //Types.Add(new CellContainsFormat());
                    break;
                case CFRuleType.Ranked:
                    Types.Add(new RankedFormat());
                    break;
                case CFRuleType.Average:
                    //Types.Add(new CellContainsFormat());
                    break;
                case CFRuleType.UniqueDuplicates:
                    //Types.Add(new CellContainsFormat());
                    break;
                case CFRuleType.CustomExpression:
                    //Types.Add(new CellContainsFormat());
                    break;
            }
        }
    }
}
