using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public PowerPoint.Shape ApplyTriangleEffect(string overlayColor1, string overlayColor2, int transparency)
        {
            float width1 = SlideHeight;
            float height1 = SlideWidth;
            float centerLeft1 = SlideWidth / 2;
            float centerTop1 = SlideHeight / 2;
            // the bigger triangle
            PowerPoint.Shape triangle1 = Shapes.AddShape(MsoAutoShapeType.msoShapeIsoscelesTriangle,
                centerLeft1 - centerTop1, centerLeft1 + centerTop1 - SlideWidth, width1, height1);
            triangle1.Rotation = 90;
            ChangeName(triangle1, EffectName.Overlay);
            triangle1.Fill.Solid();
            triangle1.Fill.ForeColor.RGB = GraphicsUtil.ConvertColorToRgb(StringUtil.GetColorFromHexValue(overlayColor1));
            triangle1.Fill.Transparency = (float)transparency / 100;
            triangle1.Line.Visible = MsoTriState.msoFalse;

            float width2 = SlideHeight / 2;
            float height2 = SlideWidth / 2;
            float centerLeft2 = SlideWidth / 4 * 3;
            float centerTop2 = SlideHeight / 4 * 3;
            // the smaller triangle
            PowerPoint.Shape triangle2 = Shapes.AddShape(MsoAutoShapeType.msoShapeIsoscelesTriangle,
                centerLeft2 + centerTop2 - SlideHeight,
                centerTop2 + SlideWidth / 2 - centerLeft2,
                width2,
                height2);
            triangle2.Rotation = 270;
            ChangeName(triangle2, EffectName.Overlay);
            triangle2.Fill.Solid();
            triangle2.Fill.ForeColor.RGB = GraphicsUtil.ConvertColorToRgb(StringUtil.GetColorFromHexValue(overlayColor2));
            triangle2.Fill.Transparency = (float)transparency / 100;
            triangle2.Line.Visible = MsoTriState.msoFalse;

            PowerPoint.Shape result = Shapes.Range(new[] { triangle1.Name, triangle2.Name }).SafeGroup(this);
            ChangeName(result, EffectName.Overlay);
            return result;
        }
    }
}
