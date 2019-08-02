using System;
using System.Drawing;

using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;
using ThreeDFormat = Microsoft.Office.Interop.PowerPoint.ThreeDFormat;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class ContourColorFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync ContourColor Format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Shape shape = shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle, 0, 0, 
                    SyncFormatConstants.DisplayImageSize.Width,
                    SyncFormatConstants.DisplayImageSize.Height);
            shape.Fill.ForeColor.RGB = formatShape.ThreeD.ContourColor.RGB;
            shape.Fill.ForeColor.TintAndShade = formatShape.ThreeD.ContourColor.TintAndShade;
            
            shape.Line.Visible = MsoTriState.msoFalse;
            Bitmap image = new Bitmap(GraphicsUtil.ShapeToBitmap(shape));
            shape.SafeDelete();
            return image;
        }
        
        private static bool Sync(Shape formatShape, Shape newShape)
        {
            ThreeDFormat source = formatShape.ThreeD;
            ThreeDFormat dest = newShape.ThreeD;

            try
            {
                // do not set SchemeColor, Brightness & ObjectThemeColor, setting them throws exceptions
                dest.ContourColor.RGB = source.ContourColor.RGB;
                dest.ContourColor.TintAndShade = source.ContourColor.TintAndShade;

                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync ContourColorFormat");
                return false;
            }

        }
        

    }
}
