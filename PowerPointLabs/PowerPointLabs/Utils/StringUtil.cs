using System;
using System.Drawing;

using Microsoft.Office.Core;

namespace PowerPointLabs.Utils
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
            string hex = BitConverter.ToString(rgbArray);
            return "#" + hex.Replace("-", "");
        }

        public static Color GetColorFromHexValue(string color)
        {
            return ColorTranslator.FromHtml(color);
        }

        public static string GetTextEffectAlignment(MsoTextEffectAlignment alignment)
        {
            switch (alignment)
            {
                case MsoTextEffectAlignment.msoTextEffectAlignmentLeft:
                    return "left";
                case MsoTextEffectAlignment.msoTextEffectAlignmentCentered:
                    return "centered";
                case MsoTextEffectAlignment.msoTextEffectAlignmentRight:
                    return "right";
                case MsoTextEffectAlignment.msoTextEffectAlignmentLetterJustify:
                    return "letterJustify";
                case MsoTextEffectAlignment.msoTextEffectAlignmentStretchJustify:
                    return "stretchJustify";
                case MsoTextEffectAlignment.msoTextEffectAlignmentWordJustify:
                    return "wordJustify";
                case MsoTextEffectAlignment.msoTextEffectAlignmentMixed:
                    return "mixed";
                default:
                    return "";
            }
        }

        public static MsoTextEffectAlignment GetTextEffectAlignment(string alignment)
        {
            switch (alignment)
            {
                case "left":
                    return MsoTextEffectAlignment.msoTextEffectAlignmentLeft;
                case "centered":
                    return MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
                case "right":
                    return MsoTextEffectAlignment.msoTextEffectAlignmentRight;
                case "letterJustify":
                    return MsoTextEffectAlignment.msoTextEffectAlignmentLetterJustify;
                case "stretchJustify":
                    return MsoTextEffectAlignment.msoTextEffectAlignmentStretchJustify;
                case "wordJustify":
                    return MsoTextEffectAlignment.msoTextEffectAlignmentWordJustify;
                // case "mixed":
                default:
                    return MsoTextEffectAlignment.msoTextEffectAlignmentMixed;
            }
        }
    }
}
