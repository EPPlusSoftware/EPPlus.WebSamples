using System;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses
{
    [Flags]
    public enum CFRuleType
    {
        AllCells = 0,
        CellContains = 2,
        Ranked = 4,
        Average = 8,
        UniqueDuplicates = 16,
        CustomExpression = 32,
        All = AllCells | CellContains | Ranked | Average | UniqueDuplicates | CustomExpression,
    }

    public enum Ranked
    {
        Top,
        Bottom
    }

    public enum Averages
    {
        Above,
        Below,
        Equal_Or_Above,
        Equal_Or_Below,
        One_Std_Dev_Above,
        One_Std_Dev_Below,
        Two_Std_Dev_Above,
        Two_Std_Dev_Below,
        Three_Std_Dev_Above,
        Three_Std_Dev_Below,
    }

    public enum UniqueDuplicate
    {
        Duplicate,
        Unique
    }

    public enum CellValueCondition
    {
        Between,
        Not_Between,
        Equal_To,
        Not_Equal_To,
        Greater_Than,
        Less_Than,
        Greater_Than_Or_Equal_To,
        Less_Than_Or_Equal_To
    }

    public enum CellContains
    {
        Cell_Value,
        Specific_Text,
        Dates_Occuring,
        Blanks,
        No_Blanks,
        Errors,
        No_Errors
    }
    public enum SpecificTextCondition
    {
        Containing,
        Not_Containing,
        Beginning_With,
        Ending_With
    }
    public enum DateCondition
    {
        Yesterday,
        Today,
        Tomorrow,
        Last_7_Days,
        Last_Week,
        This_Week,
        Next_Week,
        Last_Month,
        This_Month,
        Next_Month
    }
}
