﻿using Microsoft.Office.Core;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public PowerPoint.Shape ApplyAlbumFrameEffect(string overlayColor, int transparency)
        {
            var halfFrameWidth = 15;
            var width = SlideWidth - halfFrameWidth * 2;
            var height = SlideHeight - halfFrameWidth * 2;
            var frameShape = Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, halfFrameWidth, halfFrameWidth,
                width, height);
            ChangeName(frameShape, EffectName.Overlay);
            frameShape.Fill.Transparency = 1f;
            frameShape.Line.ForeColor.RGB = GraphicsUtil.ConvertColorToRgb(StringUtil.GetColorFromHexValue(overlayColor));
            frameShape.Line.Transparency = (float)transparency / 100;
            frameShape.Line.Weight = 30;
            frameShape.Line.Visible = MsoTriState.msoTrue;
            return frameShape;
        }
    }
}
