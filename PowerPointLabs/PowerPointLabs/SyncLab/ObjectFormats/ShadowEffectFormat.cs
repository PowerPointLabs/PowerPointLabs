using System.Drawing;

using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
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
                // setting Type to MixedType throws an error, skip it 
                // MixedType requires manual configuration of each shadow setting
                // we are unable to figure out of a shape has a custom perspective shadow or custom outer shadow
                // see MightHaveCustomPerspectiveShadow(..) for more information
            
                // setting style sets RotateWithShape, skip RotateWithShape
                // setting style sets Obscured, skip Obscured
                destFormat.Style = srcFormat.Style;
                destFormat.ForeColor = srcFormat.ForeColor;
                destFormat.Blur = srcFormat.Blur;
                destFormat.OffsetX = srcFormat.OffsetX;
                destFormat.OffsetY = srcFormat.OffsetY;
                destFormat.Size = srcFormat.Size;
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
            duplicate.SafeDelete();

            return image;

        }
        
        public static bool MightHaveCustomPerspectiveShadow(Shape shape)
        {
            // we are unable to figure out of a shape has a custom perspective shadow or custom outer shadow
            // also, we are unable to give a shape a custom perspective shadow.
            // there are 5 perspective shadow types, but the api does not tell us which was used
            return shape.Shadow.Type == MsoShadowType.msoShadowMixed &&
                   (shape.Shadow.Style == MsoShadowStyle.msoShadowStyleMixed ||
                    shape.Shadow.Style == MsoShadowStyle.msoShadowStyleOuterShadow);
        }

    }
}