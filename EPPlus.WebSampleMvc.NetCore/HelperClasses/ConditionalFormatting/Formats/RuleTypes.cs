using System;
using System.Collections.Generic;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormatting.Formats
{
    public class RuleTypes
    {
        //Dictionary<CFRuleType, IExcelConditionalFormattingHTMLRuleGroup> Types = new Dictionary<CFRuleType, IExcelConditionalFormattingHTMLRuleGroup>();
        public List<IExcelConditionalFormattingHTMLRuleGroup> Types;

        public IExcelConditionalFormattingHTMLRuleGroup ActiveRule;

        public int ActiveIndex = 0;
        //Dictionary<CFRuleType, IExcelConditionalFormattingHTMLRuleGroup?> dict = new Dictionary<CFRuleType, IExcelConditionalFormattingHTMLRuleGroup?>()
        //{
        //    { CFRuleType.CellContains, null},
        //};

        public RuleTypes()
        {
            Types = new List<IExcelConditionalFormattingHTMLRuleGroup>();
            foreach (var item in Enum.GetValues(typeof(CFRuleCategory)))
            {
                InitializeTypes((CFRuleCategory)item);
            }
            ActiveRule = Types[0];
        }

        public RuleTypes(CFRuleCategory ruleType, string[] selectedEnums = null, string[] formulas = null, FormatSettings settings = null, bool? checkbox = null)
        {
            Types = new List<IExcelConditionalFormattingHTMLRuleGroup>();
            var input = new InputData(selectedEnums, formulas, settings, checkbox);
            //int inListOrder = -1;

            switch (ruleType)
            {
                case CFRuleCategory.AllCells:
                    ActiveIndex = 0;
                    break;
                case CFRuleCategory.CellContains:
                    ActiveRule = new CellContainsFormat(input);
                    ActiveIndex = 0;
                    break;
                case CFRuleCategory.Ranked:
                    ActiveRule = new RankedFormat(input);
                    ActiveIndex = 1;
                    break;
                case CFRuleCategory.Average:
                    ActiveRule = new AverageFormat(input);
                    ActiveIndex = 2;
                    break;
                case CFRuleCategory.UniqueDuplicates:
                    ActiveRule = new UniqueDuplicateFormat(input);
                    ActiveIndex = 3;
                    break;
                case CFRuleCategory.CustomExpression:
                    ActiveRule = new ExpressionFormat(input);
                    ActiveIndex = 4;
                    break;
            }

            //ActiveRule = Types[0];

            foreach (var item in Enum.GetValues(typeof(CFRuleCategory)))
            {
                if ((CFRuleCategory)item != ruleType)
                {
                    InitializeTypes((CFRuleCategory)item);
                }
            }

            Types.Insert(ActiveIndex, ActiveRule);
        }

        void InitializeTypes(CFRuleCategory ruleType)
        {
            switch (ruleType)
            {
                case CFRuleCategory.AllCells:
                    //Types.Add(new CellContainsFormat());
                    break;
                case CFRuleCategory.CellContains:
                    Types.Add(new CellContainsFormat());
                    break;
                case CFRuleCategory.Ranked:
                    Types.Add(new RankedFormat());
                    break;
                case CFRuleCategory.Average:
                    Types.Add(new AverageFormat());
                    break;
                case CFRuleCategory.UniqueDuplicates:
                    Types.Add(new UniqueDuplicateFormat());
                    break;
                case CFRuleCategory.CustomExpression:
                    Types.Add(new ExpressionFormat());
                    break;
            }
        }
    }
}
