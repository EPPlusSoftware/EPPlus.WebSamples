using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.ContentDescriptors;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew.AdvancedRules
{
    public enum Fills
    {
        Solid_Fill,
        Gradient_Fill
    }

    public enum Borders
    {
        No_Border,
        Solid_Border
    }

    public enum BarDirecton
    {
        Context,
        LeftToRight,
        RightToLeft,
    }

    public enum AxisPosition
    {
        Automatic,
        Cell_MidPoint,
        None,
    }

    public class DataBar
    {
        public bool ShowBarOnly;

        public TypeFormula<DataBarTypes> Minimum, Maximum;

        public TypeColor<Fills> PositiveFill, NegativeFill;

        public TypeColor<Borders> PositiveBorderFill, NegativeBorderFill;

        public DropDown<BarDirecton> BarDirection;

        public TypeColor<AxisPosition> Axis;

        public DataBar()
        {
            Minimum = new TypeFormula<DataBarTypes>(DataBarTypes.Automatic);
            Maximum = new TypeFormula<DataBarTypes>(DataBarTypes.Automatic);

            PositiveFill = new TypeColor<Fills>(Fills.Solid_Fill);
            NegativeFill = new TypeColor<Fills>(PositiveFill.Selected, Color.Red);

            PositiveBorderFill = new TypeColor<Borders>(Borders.No_Border, Color.Black);
            NegativeBorderFill = new TypeColor<Borders>(PositiveBorderFill.Selected, Color.Black);

            BarDirection = new DropDown<BarDirecton>(BarDirecton.Context);

            Axis = new TypeColor<AxisPosition>(AxisPosition.Automatic, Color.Black);
        }

        //public override ExcelConditionalFormattingRule GetCFForRange(IRangeConditionalFormatting targetRange)
        //{

        //}
    }
}
