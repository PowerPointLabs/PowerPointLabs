using System;
using System.Drawing;

using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class GlowColorFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            GlowFormat glow = formatShape.Glow;
            return glow.Radius > 0 && glow.Transparency > 0.0f;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync GlowEffect Format");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Shape shape = shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0,
                    SyncFormatConstants.DisplayImageSize.Width,
                    SyncFormatConstants.DisplayImageSize.Height);
            shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            
            shape.Fill.ForeColor.RGB = formatShape.Glow.Color.RGB;
            shape.Fill.ForeColor.Brightness = formatShape.Glow.Color.Brightness;
            shape.Fill.ForeColor.TintAndShade = formatShape.Glow.Color.TintAndShade;
            shape.Fill.Solid();
            
            Bitmap image = GraphicsUtil.ShapeToBitmap(shape);
            shape.SafeDelete();
            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                GlowFormat dest = newShape.Glow;
                GlowFormat source = formatShape.Glow;

                // Color.SchemeColor must be skipped, setting it sometimes throws an exception for unknown reasons.
                // Color.ObjectThemeColor must be set despite the unrelated description in documentation.
                // The color intensity of glow will not match otherwise
                
                // syncing NotFavoriteColor throws an exception
                if (source.Color.ObjectThemeColor != MsoThemeColorIndex.msoNotThemeColor)
                {
                    dest.Color.ObjectThemeColor = source.Color.ObjectThemeColor;
                }
                
                dest.Color.RGB = source.Color.RGB;
                dest.Color.Brightness = source.Color.Brightness;
                dest.Color.TintAndShade = source.Color.TintAndShade;
                return true;
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync GlowColorFormat");
                return false;
            }
        }
    }
}
