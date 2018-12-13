using Microsoft.Office.Core;

using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.Utils;

namespace PowerPointLabs.PictureSlidesLab.Service
{
    partial class EffectsDesigner
    {
        public void ApplyTextboxEffect(string overlayColor, int transparency, int fontSizeToIncrease)
        {
            Microsoft.Office.Interop.PowerPoint.Shape shape = Util.ShapeUtil.GetTextShapeToProcess(Shapes);
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

                    Microsoft.Office.Interop.PowerPoint.Shape overlayShape = ApplyOverlayEffect(overlayColor, transparency,
                        left, top, width, height);
                    ChangeName(overlayShape, EffectName.TextBox);
                    Utils.ShapeUtil.MoveZToJustBehind(overlayShape, shape);
                }
            }
        }
    }
}
