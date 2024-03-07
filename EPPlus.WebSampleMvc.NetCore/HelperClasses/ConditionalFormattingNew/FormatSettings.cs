using System.Drawing;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew
{
    public class FormatSettings
    {
        public Color Color;

        public FormatSettings()
        {

        }

        public FormatSettings(CFColor color) { SetColor(color); }

        public void SetColor(CFColor inputColor)
        {
            Color = ColorHandler.GetColorAsColor(inputColor);
        }
    }
}
