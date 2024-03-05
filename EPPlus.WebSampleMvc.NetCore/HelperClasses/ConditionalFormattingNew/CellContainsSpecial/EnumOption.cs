using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.CellContainsSpecial
{
    internal abstract class CellContaingBase
    {
        public abstract eCellContains Type { get; }
    }
    internal abstract class CellWithFormulas<T> where T : Enum
    {
        public abstract eCellContains Type { get; }
        public abstract string[] Formulas { get; }
    }

    internal class CellValues : CellWithFormulas<CellValueCondition>
    {
        public override eCellContains Type => eCellContains.Cell_Value;

        public override string[] Formulas => throw new NotImplementedException();
    }


    internal abstract class CellValueBase
    {
        public abstract CellValueCondition Type { get; }
    }

    internal class CellValueBetween : CellValueBase
    {
        public override CellValueCondition Type => CellValueCondition.Between;

        public string[] Formulas { get; set; } = new string[2];
    }

    internal class CellValueNotBetween : CellValueBase
    {
        public override CellValueCondition Type => CellValueCondition.Not_Between;

        public string[] Formulas { get; set; } = new string[2];
    }




    //internal class EnumOption
    //{
    //    public EnumOption() { }
    //}
}
