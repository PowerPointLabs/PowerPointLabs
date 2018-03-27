using System;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointLabs.Utils;
using ShadowFormat = Microsoft.Office.Interop.PowerPoint.ShadowFormat;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    public class ShadowEffectFormat: Format
    {
        public override bool CanCopy(Shape formatShape)
        {
            return formatShape.Shadow.Visible.Equals(MsoTriState.msoTrue);
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            ShadowFormat srcFormat = formatShape.Shadow;
            ShadowFormat destFormat = newShape.Shadow;
            // the order in which the items are applied is extremely important
            // Type seems to change other items like Offset once it is set
            // A more through investigation is required to learn the exact effects
            destFormat.Visible = srcFormat.Visible;
            destFormat.Type = srcFormat.Type;
            destFormat.Style = srcFormat.Style;
            destFormat.ForeColor = srcFormat.ForeColor;
            destFormat.Obscured = srcFormat.Obscured;
            destFormat.RotateWithShape = srcFormat.RotateWithShape;
            destFormat.Blur = srcFormat.Blur;
            destFormat.OffsetX = srcFormat.OffsetX;
            destFormat.OffsetY = srcFormat.OffsetY;
            destFormat.Transparency = srcFormat.Transparency;
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            // change transparency to make shadow more visible
            // don't bother changing fill transparency, it does not affect picture shapes
            float oldShadowTransparency = formatShape.Shadow.Transparency;
            formatShape.Shadow.Transparency = 0.3f;

            Bitmap image = GraphicsUtil.ShapeToBitmap(formatShape);
            
            formatShape.Shadow.Transparency = oldShadowTransparency;

            return image;

        }
    }
}