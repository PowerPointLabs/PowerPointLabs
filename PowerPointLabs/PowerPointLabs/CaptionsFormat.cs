using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using System.Drawing;

namespace PowerPointLabs
{
    class CaptionsFormat
    {
        public static MsoTextEffectAlignment defaultAlignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
        public static Color defaultColor = Color.White;
        public static float defaultSize = 12;
        public static bool defaultBold = false;
        public static bool defaultItalic = false;
        public static Color defaultFillColor = Color.Black;
        public static string defaultFont = "Calibri";
    }
}
