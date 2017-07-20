﻿using System;
using Microsoft.Office.Core;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public PowerPoint.Shape ApplyCircleRingsEffect(string color, int transparency)
        {
            var innerCircleShape = ApplyCircleBannerEffect(color, transparency);
            var outerCircleShape = ApplyCircleBannerEffect(color, transparency, isOutline: true, margin: 10);
            if (innerCircleShape == null || outerCircleShape == null)
            {
                return null;
            }

            outerCircleShape.Left = innerCircleShape.Left + innerCircleShape.Width / 2 - outerCircleShape.Width / 2;
            outerCircleShape.Top = innerCircleShape.Top + innerCircleShape.Height / 2 - outerCircleShape.Height / 2;
            CropPicture(innerCircleShape);
            CropPicture(outerCircleShape);

            var result = Shapes.Range(new[] { innerCircleShape.Name, outerCircleShape.Name }).Group();
            ChangeName(result, EffectName.Overlay);
            return result;
        }

        private PowerPoint.Shape ApplyCircleBannerEffect(string overlayColor, int transparency,
            bool isOutline = false, int margin = 0)
        {
            var tbInfo =
                new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .GetTextBoxesInfo();
            if (tbInfo == null)
            {
                return null;
            }

            TextBoxes.AddMargin(tbInfo, margin);

            var overlayShape = ApplyCircleOverlayEffect(overlayColor, transparency, tbInfo.Left, tbInfo.Top, tbInfo.Width,
                tbInfo.Height, isOutline);
            ChangeName(overlayShape, EffectName.Banner);
            return overlayShape;
        }

        private PowerPoint.Shape ApplyCircleOverlayEffect(string color, int transparency,
            float left, float top, float width, float height, bool isOutline)
        {
            var radius = (float)Math.Sqrt(width * width / 4 + height * height / 4);
            var circleLeft = left - radius + width / 2;
            var circleTop = top - radius + height / 2;
            var circleWidth = radius * 2;

            var overlayShape = Shapes.AddShape(MsoAutoShapeType.msoShapeOval, circleLeft, circleTop,
                circleWidth, circleWidth);
            overlayShape.Fill.Solid();
            overlayShape.Fill.ForeColor.RGB = GraphicsUtil.ConvertColorToRgb(StringUtil.GetColorFromHexValue(color));
            overlayShape.Fill.Transparency = (float)transparency / 100;
            overlayShape.Line.ForeColor.RGB = GraphicsUtil.ConvertColorToRgb(StringUtil.GetColorFromHexValue(color));
            overlayShape.Line.Transparency = (float)transparency / 100;
            overlayShape.Line.Weight = 5;
            if (isOutline)
            {
                overlayShape.Fill.Visible = MsoTriState.msoFalse;
                overlayShape.Line.Visible = MsoTriState.msoTrue;
            }
            else
            {
                overlayShape.Fill.Visible = MsoTriState.msoTrue;
                overlayShape.Line.Visible = MsoTriState.msoFalse;
            }
            // as picture shape
            overlayShape.Cut();
            overlayShape = Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPastePNG)[1];
            overlayShape.Left = circleLeft;
            overlayShape.Top = circleTop;
            ChangeName(overlayShape, EffectName.Overlay);
            return overlayShape;
        }
    }
}
