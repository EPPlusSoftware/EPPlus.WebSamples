using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.AdvancedRules;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew
{
    internal class RuleTypeCollection
    {
        public List<object> Types;

        public object ActiveRule;

        public int ActiveIndex = 0;

        public FormatSettings Settings = new FormatSettings();

        public RuleTypeCollection()
        {
            Types = new List<object>();
            foreach (var type in Enum.GetValues(typeof(CFRuleType)))
            {
                InitializeTypes((CFRuleType)type);
            }
            ActiveRule = Types[0];
        }

        public RuleTypeCollection(CFRuleType ruleType, string[] selectedEnums = null, string[] formulas = null, FormatSettings settings = null, bool? checkbox = null)
        {
            Types = new List<object>();
            var input = new InputData(selectedEnums, formulas, settings, checkbox);
            //int inListOrder = -1;

            //switch (ruleType)
            //{
            //    case CFRuleType.AllCells:
            //        ActiveRule = AdvancedFormattingRules(input);
            //        break;
            //    case CFRuleType.CellContains:
            //        ActiveRule = new CellContainsFormat(input);
            //        ActiveIndex = 0;
            //        break;
            //    case CFRuleType.Ranked:
            //        ActiveRule = new RankedFormat(input);
            //        ActiveIndex = 1;
            //        break;
            //    case CFRuleType.Average:
            //        ActiveRule = new AverageFormat(input);
            //        ActiveIndex = 2;
            //        break;
            //    case CFRuleType.UniqueDuplicates:
            //        ActiveRule = new UniqueDuplicateFormat(input);
            //        ActiveIndex = 3;
            //        break;
            //    case CFRuleType.CustomExpression:
            //        ActiveRule = new ExpressionFormat(input);
            //        ActiveIndex = 4;
            //        break;
            //}

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
                    Types.Add(new AdvancedFormattingRules());
                    break;
                case CFRuleType.CellContains:
                    Types.Add(new CellContains());
                    break;
                case CFRuleType.Ranked:
                    Types.Add(new AverageRule());
                    break;
                case CFRuleType.Average:
                    Types.Add(new AverageRule());
                    break;
                case CFRuleType.UniqueDuplicates:
                    Types.Add(new UniqueDuplicateRule());
                    break;
                case CFRuleType.CustomExpression:
                    Types.Add(new ExpressionRule());
                    break;
            }
        }
    }
}
