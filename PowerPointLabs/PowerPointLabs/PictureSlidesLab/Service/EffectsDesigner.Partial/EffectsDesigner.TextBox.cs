using Microsoft.Office.Core;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.PictureSlidesLab.Util;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public void ApplyTextboxEffect(string overlayColor, int transparency, int fontSizeToIncrease)
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

                    var overlayShape = ApplyOverlayEffect(overlayColor, transparency,
                        left, top, width, height);
                    ChangeName(overlayShape, EffectName.TextBox);
                    Utils.ShapeUtil.MoveZToJustBehind(overlayShape, shape);
                }
            }
        }
    }
}
