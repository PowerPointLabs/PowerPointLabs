using Microsoft.Office.Core;

using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        // add overlay layer 
        public PowerPoint.Shape ApplyOverlayEffect(string color, int transparency,
            float left = 0, float top = 0, float? width = null, float? height = null)
        {
            width = width ?? SlideWidth;
            height = height ?? SlideHeight;
            PowerPoint.Shape overlayShape = Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, left, top,
                width.Value, height.Value);
            ChangeName(overlayShape, EffectName.Overlay);
            overlayShape.Fill.Solid();
            overlayShape.Fill.ForeColor.RGB = GraphicsUtil.ConvertColorToRgb(StringUtil.GetColorFromHexValue(color));
            overlayShape.Fill.Transparency = (float)transparency / 100;
            overlayShape.Line.ForeColor.RGB = GraphicsUtil.ConvertColorToRgb(StringUtil.GetColorFromHexValue(color));
            overlayShape.Line.Transparency = (float)transparency / 100;
            overlayShape.Line.Weight = 5;
            overlayShape.Line.Visible = MsoTriState.msoFalse;
            return overlayShape;
        }
    }
}
