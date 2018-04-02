using System;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class GlowEffectFormat: Format
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
            // emphasize the glow radius & transparency in the image preview
            float oldRadius = formatShape.Glow.Radius;
            float min = Math.Min(formatShape.Height, formatShape.Width);
            formatShape.Glow.Radius = (float) (min * 0.6);

            float threshold = 0.5f;
            float oldTransparency = formatShape.Glow.Transparency;
            if (oldTransparency < threshold)
            {
                formatShape.Glow.Transparency = 0.5f;
            }
            
            Bitmap image = GraphicsUtil.ShapeToBitmap(formatShape);
            formatShape.Glow.Radius = oldRadius;
            formatShape.Glow.Transparency = oldTransparency;
            
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
                dest.Color.ObjectThemeColor = source.Color.ObjectThemeColor;
                dest.Color.RGB = source.Color.RGB;
                dest.Color.Brightness = source.Color.Brightness;
                dest.Color.TintAndShade = source.Color.TintAndShade;

                dest.Transparency = source.Transparency;
                dest.Radius = source.Radius;
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
    }
}
