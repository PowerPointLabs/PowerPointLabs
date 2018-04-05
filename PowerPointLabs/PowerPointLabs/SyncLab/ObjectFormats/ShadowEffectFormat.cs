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
            
            destFormat.Visible = srcFormat.Visible;
            if (srcFormat.Type != MsoShadowType.msoShadowMixed)
            {
                destFormat.Type = srcFormat.Type;
                
                // only set ForeColor manually,
                // setting non-mixed types automatically sets other shadow settings to the shape
                destFormat.ForeColor = srcFormat.ForeColor;
            }
            else
            {
                // setting ShadowFormat to MixedType throws an error, skip it here
                // mixed type requires manual configuration of each shadow setting
                destFormat.Style = srcFormat.Style;
                destFormat.ForeColor = srcFormat.ForeColor;
                destFormat.Obscured = srcFormat.Obscured;
                destFormat.RotateWithShape = srcFormat.RotateWithShape;
                destFormat.Blur = srcFormat.Blur;
                destFormat.Size = srcFormat.Size;
                destFormat.OffsetX = srcFormat.OffsetX;
                destFormat.OffsetY = srcFormat.OffsetY;
                destFormat.Transparency = srcFormat.Transparency;
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            // change transparency to make shadow more visible in preview
            // don't bother changing fill transparency, it does not affect picture type shapes
            
            // setting transparency will change the ShadowFormat.Type to MixedType
            // use a duplicate to avoid complex control flow required for restoring ShadowFormat.Type
            Shape duplicate = formatShape.Duplicate()[1];
            duplicate.Shadow.Transparency = 0.3f;

            Bitmap image = GraphicsUtil.ShapeToBitmap(duplicate);
            duplicate.Delete();

            return image;

        }
    }
}