using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace PowerPointLabs
{
    public class HSL
    {
        public HSL()
        {
            _h = 0;
            _s = 0;
            _l = 0;
        }

        double _h;
        double _s;
        double _l;

        public double H
        {
            get { return _h; }
            set
            {
                _h = value;
                _h = _h > 1 ? 1 : _h < 0 ? 0 : _h;
            }
        }

        public double S
        {
            get { return _s; }
            set
            {
                _s = value;
                _s = _s > 1 ? 1 : _s < 0 ? 0 : _s;
            }
        }

        public double L
        {
            get { return _l; }
            set
            {
                _l = value;
                _l = _l > 1 ? 1 : _l < 0 ? 0 : _l;
            }
        }
    }

    class ColorHelper
    {
        public static Color ColorFromAhsb(int a, float h, float s, float b)
        {
            if (0 == s)
            {
                return Color.FromArgb(a, Convert.ToInt32(b * 255),
                  Convert.ToInt32(b * 255), Convert.ToInt32(b * 255));
            }

            float fMax, fMid, fMin;
            int iSextant, iMax, iMid, iMin;

            if (0.5 < b)
            {
                fMax = b - (b * s) + s;
                fMin = b + (b * s) - s;
            }
            else
            {
                fMax = b + (b * s);
                fMin = b - (b * s);
            }

            iSextant = (int)Math.Floor(h / 60f);
            if (300f <= h)
            {
                h -= 360f;
            }
            h /= 60f;
            h -= 2f * (float)Math.Floor(((iSextant + 1f) % 6f) / 2f);
            if (0 == iSextant % 2)
            {
                fMid = h * (fMax - fMin) + fMin;
            }
            else
            {
                fMid = fMin - h * (fMax - fMin);
            }

            iMax = Convert.ToInt32(fMax * 255);
            iMid = Convert.ToInt32(fMid * 255);
            iMin = Convert.ToInt32(fMin * 255);

            switch (iSextant)
            {
                case 1:
                    return Color.FromArgb(a, iMid, iMax, iMin);
                case 2:
                    return Color.FromArgb(a, iMin, iMax, iMid);
                case 3:
                    return Color.FromArgb(a, iMin, iMid, iMax);
                case 4:
                    return Color.FromArgb(a, iMid, iMin, iMax);
                case 5:
                    return Color.FromArgb(a, iMax, iMin, iMid);
                default:
                    return Color.FromArgb(a, iMax, iMid, iMin);
            }
        }

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

        public static Color GetComplementaryColor(Color originalColor)
        {
            return GetColorShiftedByAngle(originalColor, 180.0f);
        }

        public static List<Color> GetAnalogousColorsForColor(Color originalColor)
        {
            List<Color> analogousColors = new List<Color>();

            analogousColors.Add(GetColorShiftedByAngle(originalColor, -30.0f));
            analogousColors.Add(GetColorShiftedByAngle(originalColor, 30.0f));

            return analogousColors;
        }

        public static List<Color> GetTriadicColorsForColor(Color originalColor)
        {
            List<Color> triadicColors = new List<Color>();

            triadicColors.Add(GetColorShiftedByAngle(originalColor, -120.0f));
            triadicColors.Add(GetColorShiftedByAngle(originalColor, 120.0f));

            return triadicColors;
        }

        public static List<Color> GetTetradicColorsForColor(Color originalColor)
        {
            List<Color> tetradicColors = new List<Color>();

            tetradicColors.Add(GetColorShiftedByAngle(originalColor, -90.0f));
            tetradicColors.Add(GetColorShiftedByAngle(originalColor, 90.0f));
            tetradicColors.Add(GetComplementaryColor(originalColor));

            return tetradicColors;
        }

        public static List<Color> GetSplitComplementaryColorsForColor(Color originalColor)
        {
            List<Color> splitComplementaryColors = new List<Color>();

            splitComplementaryColors.Add(GetColorShiftedByAngle(originalColor, 150.0f));
            splitComplementaryColors.Add(GetColorShiftedByAngle(originalColor, 210.0f));

            return splitComplementaryColors;
        }
        /// <summary>
      /// Sets the absolute brightness of a colour
      /// </summary>
      /// <param name="c">Original colour</param>
      /// <param name="brightness">The luminance level to impose</param>
      /// <returns>an adjusted colour</returns>
      public static  Color SetBrightness(Color c, double brightness)
      {
         HSL hsl = RGB_to_HSL(c);
         hsl.L=brightness;
         return HSL_to_RGB(hsl);
      }
      
      /// <summary>
      /// Modifies an existing brightness level
      /// </summary>
      /// <remarks>
      /// To reduce brightness use a number smaller than 1. To increase brightness use a number larger tnan 1
      /// </remarks>
      /// <param name="c">The original colour</param>
      /// <param name="brightness">The luminance delta</param>
      /// <returns>An adjusted colour</returns>
      public static  Color ModifyBrightness(Color c, double brightness)
      {
         HSL hsl = RGB_to_HSL(c);
         hsl.L*=brightness;
         return HSL_to_RGB(hsl);
      }
 
      /// <summary>
      /// Sets the absolute saturation level
      /// </summary>
      /// <remarks>Accepted values 0-1</remarks>
      /// <param name="c">An original colour</param>
      /// <param name="Saturation">The saturation value to impose</param>
      /// <returns>An adjusted colour</returns>
      public static  Color SetSaturation(Color c, double Saturation)
      {
         HSL hsl = RGB_to_HSL(c);
         hsl.S=Saturation;
         return HSL_to_RGB(hsl);
      }
      
      /// <summary>
      /// Modifies an existing Saturation level
      /// </summary>
      /// <remarks>
      /// To reduce Saturation use a number smaller than 1. To increase Saturation use a number larger tnan 1
      /// </remarks>
      /// <param name="c">The original colour</param>
      /// <param name="Saturation">The saturation delta</param>
      /// <returns>An adjusted colour</returns>
      public static  Color ModifySaturation(Color c, double Saturation)
      {
         HSL hsl = RGB_to_HSL(c);
         hsl.S*=Saturation;
         return HSL_to_RGB(hsl);
      }
 
      /// <summary>
      /// Sets the absolute Hue level
      /// </summary>
      /// <remarks>Accepted values 0-1</remarks>
      /// <param name="c">An original colour</param>
      /// <param name="Hue">The Hue value to impose</param>
      /// <returns>An adjusted colour</returns>
      public static  Color SetHue(Color c, double Hue)
      {
         HSL hsl = RGB_to_HSL(c);
         hsl.H=Hue;
         return HSL_to_RGB(hsl);
      }
      
      /// <summary>
      /// Modifies an existing Hue level
      /// </summary>
      /// <remarks>
      /// To reduce Hue use a number smaller than 1. To increase Hue use a number larger tnan 1
      /// </remarks>
      /// <param name="c">The original colour</param>
      /// <param name="Hue">The Hue delta</param>
      /// <returns>An adjusted colour</returns>
      public static  Color ModifyHue(Color c, double Hue)
      {
         HSL hsl = RGB_to_HSL(c);
         hsl.H*=Hue;
         return HSL_to_RGB(hsl);
      }
 
      /// <summary>
      /// Converts a colour from HSL to RGB
      /// </summary>
      /// <remarks>Adapted from the algoritm in Foley and Van-Dam</remarks>
      /// <param name="hsl">The HSL value</param>
      /// <returns>A Color structure containing the equivalent RGB values</returns>
      public static Color HSL_to_RGB(HSL hsl)
      {
         double r=0,g=0,b=0;
         double temp1,temp2;
 
         if(hsl.L==0)
         {
            r=g=b=0;
         }
         else
         {
            if(hsl.S==0)
            {
               r=g=b=hsl.L;
            }
            else
            {
               temp2 = ((hsl.L<=0.5) ? hsl.L*(1.0+hsl.S) : hsl.L+hsl.S-(hsl.L*hsl.S));
               temp1 = 2.0*hsl.L-temp2;
            
               double[] t3=new double[]{hsl.H+1.0/3.0,hsl.H,hsl.H-1.0/3.0};
               double[] clr=new double[]{0,0,0};
               for(int i=0;i<3;i++)
               {
                  if(t3[i]<0)
                     t3[i]+=1.0;
                  if(t3[i]>1)
                     t3[i]-=1.0;
 
                  if(6.0*t3[i] < 1.0)
                     clr[i]=temp1+(temp2-temp1)*t3[i]*6.0;
                  else if(2.0*t3[i] < 1.0)
                     clr[i]=temp2;
                  else if(3.0*t3[i] < 2.0)
                     clr[i]=(temp1+(temp2-temp1)*((2.0/3.0)-t3[i])*6.0);
                  else
                     clr[i]=temp1;
               }
               r=clr[0];
               g=clr[1];
               b=clr[2];
            }
         }
 
         return Color.FromArgb((int)(255*r),(int)(255*g),(int)(255*b));
 
      }
 
 
      //
      /// <summary>
      /// Converts RGB to HSL
      /// </summary>
      /// <remarks>Takes advantage of whats already built in to .NET by using the Color.GetHue, Color.GetSaturation and Color.GetBrightness methods</remarks>
      /// <param name="c">A Color to convert</param>
      /// <returns>An HSL value</returns>
      public static HSL RGB_to_HSL (Color c)
      {
         HSL hsl =  new HSL();
         
         hsl.H=c.GetHue()/360.0; // we store hue as 0-1 as opposed to 0-360
         hsl.L=c.GetBrightness();
         hsl.S=c.GetSaturation();
 
         return hsl;
      }
    }
}
