using System.Collections.Generic;
using System.Drawing;
using System.Dynamic;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses
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

    public static class ColorHandler
    {
        static Dictionary<CFColor, int[]> colors = new Dictionary<CFColor, int[]>()
        {
            { CFColor.Blue, CreateArr(68, 114, 196) },
            { CFColor.Orange,  CreateArr(237, 125, 49) },
            { CFColor.Gray,  CreateArr(165, 165, 165) },
            { CFColor.Gold,  CreateArr(255, 192, 0) },
            { CFColor.Green,  CreateArr(112, 173, 71) },
            { CFColor.Red,   CreateArr(192, 0, 0) },
        };

        public static int[] GetRGBValues(CFColor color)
        {
            return colors[color];
        }

        public static Color GetColorAsColor(CFColor color)
        {
            var rgb = GetRGBValues(color);
            return Color.FromArgb(rgb[0], rgb[1], rgb[2]);
        }

        private static int[] CreateArr(int r, int g, int b)
        {
            return new int[] { r, g, b };
        }
    }
}
