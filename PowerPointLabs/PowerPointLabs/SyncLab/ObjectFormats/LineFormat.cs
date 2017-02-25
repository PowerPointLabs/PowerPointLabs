using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {

        }

        public static Bitmap DisplayImage(Shape formatShape)
        {
            return null;
        }
    }
}
