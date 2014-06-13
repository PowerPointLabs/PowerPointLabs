using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace PowerPointLabs
{
    class ColorHelper
    {
        public static Color GetColorShiftedByAngle(Color originalColor, float angle)
        {
            if (angle < 0)
            {
                while (angle < 0)
                {
                    angle += 360.0f;
                }
            }

            float baseAngle = originalColor.GetHue();
            float finalAngle = baseAngle + (angle % 360);
            
            if (finalAngle > 360.0f)
            {
                finalAngle -= 360.0f;
            }

            Color finalColor = ColorFromHSB(finalAngle, originalColor.GetSaturation(), originalColor.GetBrightness());

            return finalColor;
        }

        public static Color ColorFromHSB(float hue, float saturation, float value)
        {
            int hi = Convert.ToInt32(Math.Floor(hue / 60)) % 6;
            double f = hue / 60 - Math.Floor(hue / 60);

            value = value * 255;
            int v = Convert.ToInt32(value);
            int p = Convert.ToInt32(value * (1 - saturation));
            int q = Convert.ToInt32(value * (1 - f * saturation));
            int t = Convert.ToInt32(value * (1 - (1 - f) * saturation));

            switch (hi)
            {
                case 0: return Color.FromArgb(255, v, t, p);
                case 1: return Color.FromArgb(255, q, v, p);
                case 2: return Color.FromArgb(255, p, v, t); 
                case 3: return Color.FromArgb(255, p, q, v);
                case 4: return Color.FromArgb(255, t, p, v);
                default: return Color.FromArgb(255, v, p, q); 
            }
        }

        public Color GetComplementaryColor(Color originalColor)
        {
            return GetColorShiftedByAngle(originalColor, 180.0f);
        }

        public List<Color> GetAnalogousColorsForColor(Color originalColor)
        {
            List<Color> analogousColors = new List<Color>();

            analogousColors.Add(GetColorShiftedByAngle(originalColor, -30.0f));
            analogousColors.Add(GetColorShiftedByAngle(originalColor, 30.0f));

            return analogousColors;
        }

        public List<Color> GetTriadicColorsForColor(Color originalColor)
        {
            List<Color> triadicColors = new List<Color>();

            triadicColors.Add(GetColorShiftedByAngle(originalColor, -120.0f));
            triadicColors.Add(GetColorShiftedByAngle(originalColor, 120.0f));

            return triadicColors;
        }

        public List<Color> GetTetradicColorsForColor(Color originalColor)
        {
            List<Color> tetradicColors = new List<Color>();

            tetradicColors.Add(GetColorShiftedByAngle(originalColor, -90.0f));
            tetradicColors.Add(GetColorShiftedByAngle(originalColor, 90.0f));
            tetradicColors.Add(GetComplementaryColor(originalColor));

            return tetradicColors;
        }

        public List<Color> GetSplitComplementaryColorsForColor(Color originalColor)
        {
            List<Color> splitComplementaryColors = new List<Color>();

            splitComplementaryColors.Add(GetColorShiftedByAngle(originalColor, 150.0f));
            splitComplementaryColors.Add(GetColorShiftedByAngle(originalColor, 210.0f));

            return splitComplementaryColors;
        }
    }
}
