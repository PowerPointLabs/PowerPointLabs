using Microsoft.Office.Core;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public void ApplyTextboxEffect(string overlayColor, int transparency)
        {
            foreach (PowerPoint.Shape shape in Shapes)
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
                        var left = paragraph.BoundLeft - 5;
                        var top = paragraph.BoundTop;
                        var width = paragraph.BoundWidth + 10;
                        var height = paragraph.BoundHeight;

                        var overlayShape = ApplyOverlayEffect(overlayColor, transparency,
                            left, top, width, height);
                        ChangeName(overlayShape, EffectName.TextBox);
                        Graphics.MoveZToJustBehind(overlayShape, shape);
                        shape.Tags.Add(Tag.AddedTextbox, overlayShape.Name);
                    }
                }
            }
            foreach (PowerPoint.Shape shape in Shapes)
            {
                shape.Tags.Add(Tag.AddedTextbox, "");
            }
        }
    }
}
