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
            //int inListOrder = -1;

            switch (ruleType)
            {
                case CFRuleType.AllCells:
                    ActiveIndex = 0;
                    break;
                case CFRuleType.CellContains:
                    ActiveRule = new CellContainsFormat(input);
                    ActiveIndex = 0;
                    break;
                case CFRuleType.Ranked:
                    ActiveRule = new RankedFormat(input);
                    ActiveIndex = 1;
                    break;
                case CFRuleType.Average:
                    ActiveRule = new AverageFormat(input);
                    ActiveIndex = 2;
                    break;
                case CFRuleType.UniqueDuplicates:
                    ActiveRule = new UniqueDuplicateFormat(input);
                    ActiveIndex = 3;
                    break;
                case CFRuleType.CustomExpression:
                    ActiveRule = new ExpressionFormat(input);
                    ActiveIndex = 4;
                    break;
            }

            //ActiveRule = Types[0];

            foreach (var item in Enum.GetValues(typeof(CFRuleType)))
            {
                if ((CFRuleType)item != ruleType)
                {
                    InitializeTypes((CFRuleType)item);
                }
            }

            Types.Insert(ActiveIndex, ActiveRule);
        }

        void InitializeTypes(CFRuleType ruleType)
        {
            switch (ruleType)
            {
                case CFRuleType.AllCells:
                    //Types.Add(new CellContainsFormat());
                    break;
                case CFRuleType.CellContains:
                    Types.Add(new CellContainsFormat());
                    break;
                case CFRuleType.Ranked:
                    Types.Add(new RankedFormat());
                    break;
                case CFRuleType.Average:
                    Types.Add(new AverageFormat());
                    break;
                case CFRuleType.UniqueDuplicates:
                    Types.Add(new UniqueDuplicateFormat());
                    break;
                case CFRuleType.CustomExpression:
                    Types.Add(new ExpressionFormat());
                    break;
            }
        }
    }
}
