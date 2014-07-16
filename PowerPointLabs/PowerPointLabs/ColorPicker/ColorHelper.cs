using System.Collections.Generic;

namespace PowerPointLabs.ColorPicker
{
    class ColorHelper
    {
        public static int ReverseRGBToArgb(int x)
        {
            int R = 0xff & x;
            int G = (0xff00 & x) >> 8;
            int B = (0xff0000 & x) >> 16;
            return (int)(0xff << 24 | R << 16 | G << 8 | B);
        }

        public static HSLColor GetColorShiftedByAngle(HSLColor originalColor, float angle)
        {
            if (angle < 0)
            {
                while (angle < 0)
                {
                    angle += 360.0f;
                }
            }

            var baseAngle = (float) originalColor.Hue;
            var finalAngle = baseAngle + (angle % 360);
            
            if (finalAngle > 360.0f)
            {
                finalAngle -= 360.0f;
            }
            var finalColor = new HSLColor(finalAngle, originalColor.Saturation, originalColor.Luminosity);

            return finalColor;
        }

        public static HSLColor GetComplementaryColor(HSLColor originalColor)
        {
            return GetColorShiftedByAngle(originalColor, 180.0f);
        }

        public static List<HSLColor> GetAnalogousColorsForColor(HSLColor originalColor)
        {
            var analogousColors = new List<HSLColor>
            {
                GetColorShiftedByAngle(originalColor, -30.0f),
                GetColorShiftedByAngle(originalColor, 30.0f)
            };

            return analogousColors;
        }

        public static List<HSLColor> GetTriadicColorsForColor(HSLColor originalColor)
        {
            var triadicColors = new List<HSLColor>
            {
                GetColorShiftedByAngle(originalColor, -120.0f),
                GetColorShiftedByAngle(originalColor, 120.0f)
            };

            return triadicColors;
        }

        public static List<HSLColor> GetTetradicColorsForColor(HSLColor originalColor)
        {
            var tetradicColors = new List<HSLColor>
            {
                GetColorShiftedByAngle(originalColor, -90.0f),
                GetColorShiftedByAngle(originalColor, 90.0f),
                GetComplementaryColor(originalColor)
            };

            return tetradicColors;
        }

        public static List<HSLColor> GetSplitComplementaryColorsForColor(HSLColor originalColor)
        {
            var splitComplementaryColors = new List<HSLColor>
            {
                GetColorShiftedByAngle(originalColor, 150.0f),
                GetColorShiftedByAngle(originalColor, 210.0f)
            };

            return splitComplementaryColors;
        }
    }
}
