using Microsoft.Office.Core;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public void ApplyFrostedGlassTextBoxEffect(string overlayColor, int transparency, Shape blurImage)
        {
            foreach (Shape shape in Shapes)
            {
                if ((shape.Type != MsoShapeType.msoPlaceholder
                        && shape.Type != MsoShapeType.msoTextBox)
                        || shape.TextFrame.HasText == MsoTriState.msoFalse
                        || StringUtil.IsNotEmpty(shape.Tags[Tag.AddedTextbox])
                        || StringUtil.IsNotEmpty(shape.Tags[Tag.ImageReference]))
                {
                    continue;
                }

                // multiple paragraphs.. 
                foreach (TextRange2 textRange in shape.TextFrame2.TextRange.Paragraphs)
                {
                    if (StringUtil.IsNotEmpty(textRange.TrimText().Text))
                    {
                        var paragraph = textRange.TrimText();
                        var left = paragraph.BoundLeft - 10;
                        var top = paragraph.BoundTop - 10;
                        var width = paragraph.BoundWidth + 20;
                        var height = paragraph.BoundHeight + 20;

                        blurImage.Copy();
                        var blurTextBox = Shapes.Paste()[1];
                        blurTextBox.Left = blurImage.Left;
                        blurTextBox.Top = blurImage.Top;
                        CropPicture(blurTextBox, left, top, width, height);
                        ChangeName(blurTextBox, EffectName.TextBox);

                        var overlayShape = ApplyOverlayEffect(overlayColor, transparency,
                            left, top, width, height);
                        ChangeName(overlayShape, EffectName.TextBox);

                        Graphics.MoveZToJustBehind(blurTextBox, shape);
                        Graphics.MoveZToJustBehind(overlayShape, shape);
                        shape.Tags.Add(Tag.AddedTextbox, overlayShape.Name);
                    }
                }
            }
            foreach (Shape shape in Shapes)
            {
                shape.Tags.Add(Tag.AddedTextbox, "");
            }
        }

        public Shape ApplyFrostedGlassBannerEffect(BannerDirection direction, Position textPos, Shape blurImage,
            string overlayColor, int transparency)
        {
            var tbInfo =
                new TextBoxes(Shapes.Range(), SlideWidth, SlideHeight)
                .GetTextBoxesInfo();
            if (tbInfo == null)
                return null;

            TextBoxes.AddMargin(tbInfo);

            Shape overlayShape;
            blurImage.Copy();
            Shape blurBanner = Shapes.Paste()[1];
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
            Graphics.MoveZToJustBehind(blurBanner, overlayShape);

            var range = Shapes.Range(new[] {blurBanner.Name, overlayShape.Name});
            var resultShape = range.Group();
            ChangeName(resultShape, EffectName.Banner);
            return resultShape;
        }
    }
}
