using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;


namespace PowerPointLabs.ResizeLab
{
    internal static partial class ResizeLabMain
    {
        #region API

        public static void StretchLeft(PowerPoint.ShapeRange stretchShapes)
        {
            if (stretchShapes.Count < 2)
            {
                return;
            }
            Shape a  = stretchShapes[1];
        }

        #endregion
    }
}
