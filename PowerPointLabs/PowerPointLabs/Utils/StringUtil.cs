using System;
using System.Drawing;

namespace PowerPointLabs.ImageSearch.Util
{
    class StringUtil
    {
        public static bool IsEmpty(string str)
        {
            return str == null || str.Trim().Length == 0;
        }

        public static bool IsNotEmpty(string str)
        {
            return str != null && str.Trim().Length > 0;
        }

        public static string GetHexValue(Color color)
        {
            byte[] rgbArray = { color.R, color.G, color.B };
            var hex = BitConverter.ToString(rgbArray);
            return "#" + hex.Replace("-", "");
        }
    }
}
