using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.Utils;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public void ApplyFrostedGlassTextBoxEffect(string overlayColor, int transparency, Shape blurImage, int fontSizeToIncrease)
        {
            Shape shape = Util.ShapeUtil.GetTextShapeToProcess(Shapes);
            if (shape == null)
            {
                return;
            }

            int margin = CalculateTextBoxMargin(fontSizeToIncrease);
            // multiple paragraphs.. 
            foreach (TextRange2 textRange in shape.TextFrame2.TextRange.Paragraphs)
            {
                if (StringUtil.IsNotEmpty(textRange.TrimText().Text))
                {
                    TextRange2 paragraph = textRange.TrimText();
                    float left = paragraph.BoundLeft - margin;
                    float top = paragraph.BoundTop - margin;
                    float width = paragraph.BoundWidth + margin * 2;
                    float height = paragraph.BoundHeight + margin * 2;

                    Shape blurTextBox = blurImage.Duplicate()[1];
                    blurTextBox.Left = blurImage.Left;
                    blurTextBox.Top = blurImage.Top;
                    CropPicture(blurTextBox, left, top, width, height);
                    ChangeName(blurTextBox, EffectName.TextBox);

                    Shape overlayShape = ApplyOverlayEffect(overlayColor, transparency,
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
            TextBoxInfo tbInfo =
                new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .GetTextBoxesInfo();
            if (tbInfo == null)
            {
                return null;
            }

            TextBoxes.AddMargin(tbInfo);

            Shape overlayShape;
            Shape blurBanner = blurImage.Duplicate()[1];
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

            Microsoft.Office.Interop.PowerPoint.ShapeRange range = Shapes.Range(new[] {blurBanner.Name, overlayShape.Name});
            Shape resultShape = range.SafeGroup(this);
            ChangeName(resultShape, EffectName.Banner);
            return resultShape;
        }
    }
}
