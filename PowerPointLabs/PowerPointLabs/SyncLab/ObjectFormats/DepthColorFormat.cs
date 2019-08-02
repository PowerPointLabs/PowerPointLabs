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
    class DepthColorFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return true;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync DepthColor Format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Shape shape = shapes.AddShape(
                    MsoAutoShapeType.msoShapeRectangle, 0, 0, 
                    SyncFormatConstants.DisplayImageSize.Width,
                    SyncFormatConstants.DisplayImageSize.Height);
            // do not sync extrusion theme color with forecolor, it throws an exception
            shape.Fill.ForeColor.RGB = formatShape.ThreeD.ExtrusionColor.RGB;
            shape.Fill.ForeColor.TintAndShade = formatShape.ThreeD.ExtrusionColor.TintAndShade;
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
                // don't set type if type is TypeMixed, it throws an exception
                if (source.ExtrusionColorType != MsoExtrusionColorType.msoExtrusionColorTypeMixed)
                {
                    dest.ExtrusionColorType = source.ExtrusionColorType;
                }
                if (source.ExtrusionColorType != MsoExtrusionColorType.msoExtrusionColorAutomatic)
                {
                    // do not set SchemeColor & Brightness, setting them throws exceptions
                    dest.ExtrusionColor.ObjectThemeColor = source.ExtrusionColor.ObjectThemeColor;
                    dest.ExtrusionColor.RGB = source.ExtrusionColor.RGB;
                    dest.ExtrusionColor.TintAndShade = source.ExtrusionColor.TintAndShade;
                }

                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync DepthColorFormat");
                return false;
            }

        }
        

    }
}
