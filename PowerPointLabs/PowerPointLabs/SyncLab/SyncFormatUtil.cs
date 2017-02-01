using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointLabs.SyncLab
{
    class SyncFormatUtil
    {

        public static void SyncColorFormat(ColorFormat formatToApplyOn, ColorFormat formatToApply)
        {
            try
            {
                formatToApplyOn.ObjectThemeColor = formatToApply.ObjectThemeColor;
            }
            catch (ArgumentException)
            {
            }
            try
            {
                formatToApplyOn.SchemeColor = formatToApply.SchemeColor;
            }
            catch (ArgumentException)
            {
            }
            formatToApplyOn.RGB = formatToApply.RGB;
            formatToApplyOn.Brightness = formatToApply.Brightness;
            formatToApplyOn.TintAndShade = formatToApply.TintAndShade;
        }

        public static void SyncFillFormat(FillFormat formatToApplyOn, FillFormat formatToApply)
        {
            
        }

        public static void SyncFontFormat(Font formatToApplyOn, Font formatToApply)
        {
            formatToApplyOn.AutoRotateNumbers = formatToApply.AutoRotateNumbers;
            formatToApplyOn.BaselineOffset = formatToApply.BaselineOffset;
            formatToApplyOn.Bold = formatToApply.Bold;
            formatToApplyOn.Emboss = formatToApply.Emboss;
            formatToApplyOn.Italic = formatToApply.Italic;
            formatToApplyOn.Name = formatToApply.Name;
            formatToApplyOn.NameAscii = formatToApply.NameAscii;
            formatToApplyOn.NameComplexScript = formatToApply.NameComplexScript;
            formatToApplyOn.NameFarEast = formatToApply.NameFarEast;
            formatToApplyOn.NameOther = formatToApply.NameOther;
            formatToApplyOn.Shadow = formatToApply.Shadow;
            formatToApplyOn.Size = formatToApply.Size;
            formatToApplyOn.Subscript = formatToApply.Subscript;
            formatToApplyOn.Superscript = formatToApply.Superscript;
            formatToApplyOn.Underline = formatToApply.Underline;
            SyncColorFormat(formatToApplyOn.Color, formatToApply.Color);
        }

        public static void SyncLineFormat(LineFormat formatToApplyOn, LineFormat formatToApply)
        {
            try
            {
                formatToApplyOn.BeginArrowheadLength = formatToApply.BeginArrowheadLength;
                formatToApplyOn.BeginArrowheadStyle = formatToApply.BeginArrowheadStyle;
                formatToApplyOn.BeginArrowheadWidth = formatToApply.BeginArrowheadWidth;
                formatToApplyOn.EndArrowheadLength = formatToApply.EndArrowheadLength;
                formatToApplyOn.EndArrowheadStyle = formatToApply.EndArrowheadStyle;
                formatToApplyOn.EndArrowheadWidth = formatToApply.EndArrowheadWidth;
            }
            catch (ArgumentException)
            {
            }
            formatToApplyOn.DashStyle = formatToApply.DashStyle;
            SyncColorFormat(formatToApplyOn.ForeColor, formatToApply.ForeColor);
            formatToApplyOn.InsetPen = formatToApply.InsetPen;
            try
            {
                formatToApplyOn.Pattern = formatToApply.Pattern;
            }
            catch (ArgumentException)
            {
            }
            formatToApplyOn.Style = formatToApply.Style;
            formatToApplyOn.Transparency = formatToApply.Transparency;
            formatToApplyOn.Weight = formatToApply.Weight;
        }

    }
}
