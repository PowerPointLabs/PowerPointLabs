using System.Collections.Generic;
using System.Drawing;

namespace PowerPointLabs.ColorsLab
{
    class ColorHelper
    {
        public static int ReverseRGBToArgb(int x)
        {
            int r = 0xff & x;
            int g = (0xff00 & x) >> 8;
            int b = (0xff0000 & x) >> 16;
            return (int)(0xff << 24 | r << 16 | g << 8 | b);
        }

        public static string ColorToHexString(Color color)
        {
            return "#" + color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2");
        }

        public static Color GetColorShiftedByAngle(HSLColor originalColor, float angle)
        {
            if (angle < 0)
            {
                while (angle < 0)
                {
                    angle += 360.0f;
                }
            }

            float baseAngle = (float) originalColor.Hue / 240.0f * 360.0f;
            float finalAngle = baseAngle + angle;
            
            if (finalAngle > 360.0f)
            {
                finalAngle -= 360.0f;
            }
            Color finalColor = new HSLColor(finalAngle / 360.0f * 240.0f, originalColor.Saturation, originalColor.Luminosity);

            return Color.FromArgb(255,
                    finalColor.R,
                    finalColor.G,
                    finalColor.B);
        }

        public static Color GetComplementaryColor(HSLColor originalColor)
        {
            return GetColorShiftedByAngle(originalColor, 180.0f);
        }

        public static List<Color> GetAnalogousColorsForColor(HSLColor originalColor)
        {
            List<Color> analogousColors = new List<Color>
            {
                GetColorShiftedByAngle(originalColor, -30.0f),
                GetColorShiftedByAngle(originalColor, 30.0f)
            };

            return analogousColors;
        }

        public static List<Color> GetTriadicColorsForColor(HSLColor originalColor)
        {
            List<Color> triadicColors = new List<Color>
            {
                GetColorShiftedByAngle(originalColor, -120.0f),
                GetColorShiftedByAngle(originalColor, 120.0f)
            };

            return triadicColors;
        }

        public static List<Color> GetTetradicColorsForColor(HSLColor originalColor)
        {
            List<Color> tetradicColors = new List<Color>
            {
                GetColorShiftedByAngle(originalColor, -90.0f),
                GetColorShiftedByAngle(originalColor, 90.0f),
                GetComplementaryColor(originalColor)
            };

            return tetradicColors;
        }

        public static List<Color> GetSplitComplementaryColorsForColor(HSLColor originalColor)
        {
            List<Color> splitComplementaryColors = new List<Color>
            {
                GetColorShiftedByAngle(originalColor, 150.0f),
                GetColorShiftedByAngle(originalColor, 210.0f)
            };

            return splitComplementaryColors;
        }
    }
}
