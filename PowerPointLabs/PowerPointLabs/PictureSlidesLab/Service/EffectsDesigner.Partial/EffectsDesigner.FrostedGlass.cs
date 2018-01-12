using Microsoft.Office.Core;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.PictureSlidesLab.Util;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public void ApplyFrostedGlassTextBoxEffect(string overlayColor, int transparency, Shape blurImage, int fontSizeToIncrease)
        {
            var shape = Util.ShapeUtil.GetTextShapeToProcess(Shapes);
            if (shape == null)
            {
                return;
            }

            var margin = CalculateTextBoxMargin(fontSizeToIncrease);
            // multiple paragraphs.. 
            foreach (TextRange2 textRange in shape.TextFrame2.TextRange.Paragraphs)
            {
                if (StringUtil.IsNotEmpty(textRange.TrimText().Text))
                {
                    var paragraph = textRange.TrimText();
                    var left = paragraph.BoundLeft - margin;
                    var top = paragraph.BoundTop - margin;
                    var width = paragraph.BoundWidth + margin * 2;
                    var height = paragraph.BoundHeight + margin * 2;

                    var blurTextBox = blurImage.Duplicate()[1];
                    blurTextBox.Left = blurImage.Left;
                    blurTextBox.Top = blurImage.Top;
                    CropPicture(blurTextBox, left, top, width, height);
                    ChangeName(blurTextBox, EffectName.TextBox);

                    var overlayShape = ApplyOverlayEffect(overlayColor, transparency,
                        left, top, width, height);
                    ChangeName(overlayShape, EffectName.TextBox);

                    Utils.ShapeUtil.MoveZToJustBehind(blurTextBox, shape);
                    Utils.ShapeUtil.MoveZToJustBehind(overlayShape, shape);
                }
            }
        }

        public Shape ApplyFrostedGlassBannerEffect(BannerDirection direction, Position textPos, Shape blurImage,
            string overlayColor, int transparency)
        {
            var tbInfo =
                new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .GetTextBoxesInfo();
            if (tbInfo == null)
            {
                return null;
            }

            TextBoxes.AddMargin(tbInfo);

            Shape overlayShape;
            var blurBanner = blurImage.Duplicate()[1];
            blurBanner.Left = blurImage.Left;
            blurBanner.Top = blurImage.Top;
            direction = HandleAutoDirection(direction, textPos);
            switch (direction)
            {
                case BannerDirection.Horizontal:
                    overlayShape = ApplyOverlayEffect(overlayColor, transparency, 0, tbInfo.Top, SlideWidth,
                        tbInfo.Height);
                    CropPicture(blurBanner, 0, tbInfo.Top, SlideWidth, tbInfo.Height);
                    break;
                // case BannerDirection.Vertical:
                default:
                    overlayShape = ApplyOverlayEffect(overlayColor, transparency, tbInfo.Left, 0, tbInfo.Width,
                        SlideHeight);
                    CropPicture(blurBanner, tbInfo.Left, 0, tbInfo.Width, SlideHeight);
                    break;
            }
            ChangeName(overlayShape, EffectName.Banner);
            ChangeName(blurBanner, EffectName.Banner);
            overlayShape.ZOrder(MsoZOrderCmd.msoSendToBack);
            Utils.ShapeUtil.MoveZToJustBehind(blurBanner, overlayShape);

            var range = Shapes.Range(new[] {blurBanner.Name, overlayShape.Name});
            var resultShape = range.Group();
            ChangeName(resultShape, EffectName.Banner);
            return resultShape;
        }
    }
}
