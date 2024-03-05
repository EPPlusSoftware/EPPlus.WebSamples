using System.Drawing;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew
{
    public enum CFColor
    {
        Blue,
        Orange,
        Gray,
        Gold,
        Green,
        Red
    }

    public class FormatSettings
    {
        public Color Color;

        public FormatSettings()
        {

        }

        public FormatSettings(CFColor color) { SetColor(color); }

        public void SetColor(CFColor inputColor)
        {
            //Color = ColorHandler.GetColorAsColor(inputColor);
        }
    }
}
