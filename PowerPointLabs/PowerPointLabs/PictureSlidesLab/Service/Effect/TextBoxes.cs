using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using Microsoft.Office.Core;

using PowerPointLabs.Utils;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using ShapeUtil = PowerPointLabs.PictureSlidesLab.Util.ShapeUtil;


namespace PowerPointLabs.PictureSlidesLab.Service.Effect
{
    public class TextBoxes
    {
        public const int Margin = 25;

        public const int ExtendMargin = 5;

        private List<Shape> TextShapes { get; set; }

        private readonly float _slideWidth;

        private readonly float _slideHeight;

        private Position _pos;

        private Alignment _align;

        private float _left;

        private float _top;

        # region APIs
        public TextBoxes(ShapeRange shapes, float slideWidth, float slideHeight)
        {
            _slideWidth = slideWidth;
            _slideHeight = slideHeight;
            TextShapes = new List<Shape>();
            Shape shape = ShapeUtil.GetTextShapeToProcess(shapes);
            if (shape != null)
            {
                TextShapes.Add(shape);
            }
        }

        public bool IsTextShapesEmpty()
        {
            return TextShapes.Count == 0;
        }

        public TextBoxes SetPosition(Position pos)
        {
            _pos = pos;
            return this;
        }

        public TextBoxes SetAlignment(Alignment align)
        {
            _align = align;
            return this;
        }

        public void StartBoxing()
        {
            if (_pos == Position.NoEffect)
            {
                return;
            }

            // do positioning twice to fix a bug:
            // if only do positioning once,
            // textboxes' height/top may be incorrect if each textbox is not directly next to each other;
            // doing positioning (make every textbox next to each other) the next time will fix the problem.
            StartPositioning();
            StartPositioning();
        }

        public TextBoxInfo GetTextBoxesInfo()
        {
            return TextShapes.Count > 0 ? GetTextBoxesInfo(TextShapes) : null;
        }

        public void StartTextWrapping()
        {
            foreach (Shape textShape in TextShapes)
            {
                if (textShape.Width > _slideWidth / 2)
                {
                    ShapeUtil.AddTag(textShape, Tag.OriginalShapeWidth, 
                        textShape.Width.ToString(CultureInfo.InvariantCulture));
                    textShape.Width = _slideWidth / 2;
                }
            }
        }

        public void RecoverTextWrapping()
        {
            foreach (Shape textShape in TextShapes)
            {
                if (StringUtil.IsNotEmpty(textShape.Tags[Tag.OriginalShapeWidth]))
                {
                    textShape.Width = float.Parse(textShape.Tags[Tag.OriginalShapeWidth]);
                    textShape.Tags.Add(Tag.OriginalShapeWidth, "");
                }
            }
        }

        public static void AddMargin(TextBoxInfo textboxesInfo, float? margin = null)
        {
            margin = margin ?? Margin;
            textboxesInfo.Left -= margin.Value;
            textboxesInfo.Top -= margin.Value;
            textboxesInfo.Width += 2 * margin.Value;
            textboxesInfo.Height += 2 * margin.Value;
        }

        # endregion

        # region Helper Funcs

        private void StartPositioning()
        {
            // decide which textbox is on top of the other
            SortTextBoxes();

            SetupTextBoxesAlignment();

            TextBoxInfo boxesInfo = GetTextBoxesInfo(TextShapes);
            SetupTextBoxesPosition(boxesInfo);

            float accumulatedHeight = 0f;
            foreach (Shape textShape in TextShapes)
            {
                TextBoxInfo singleBoxInfo = GetTextBoxInfo(textShape);

                AdjustShapeLeft(textShape, boxesInfo, singleBoxInfo);
                AdjustShapeTop(textShape, singleBoxInfo, accumulatedHeight);
                accumulatedHeight += singleBoxInfo.Height;
            }
        }

        private void AdjustShapeTop(Shape textShape, TextBoxInfo singleBoxInfo, float accumulatedHeight)
        {
            textShape.Top = _top + textShape.Top - (singleBoxInfo.Top - accumulatedHeight);
        }

        private void AdjustShapeLeft(Shape textShape, TextBoxInfo boxesInfo, TextBoxInfo singleBoxInfo)
        {
            switch (_align)
            {
                case Alignment.Left:
                    textShape.Left = _left + textShape.Left - (singleBoxInfo.Left - 0f);
                    break;
                case Alignment.Centre:
                    textShape.Left = _left + textShape.Left -
                                     (singleBoxInfo.Left - (boxesInfo.Width/2 - singleBoxInfo.Width/2));
                    break;
                case Alignment.Right:
                    textShape.Left = _left + textShape.Left - (singleBoxInfo.Left - (boxesInfo.Width - singleBoxInfo.Width));
                    break;
            }
        }

        private void SetupTextBoxesAlignment()
        {
            if (_align == Alignment.NoEffect)
            {
                return;
            }

            HandleAutoAlignment();
            switch (_align)
            {
                case Alignment.Left:
                    SetTextAlignment(MsoTextEffectAlignment.msoTextEffectAlignmentLeft);
                    break;
                case Alignment.Centre:
                    SetTextAlignment(MsoTextEffectAlignment.msoTextEffectAlignmentCentered);
                    break;
                case Alignment.Right:
                    SetTextAlignment(MsoTextEffectAlignment.msoTextEffectAlignmentRight);
                    break;
            }
        }

        private void HandleAutoAlignment()
        {
            if (_align != Alignment.Auto)
            {
                return;
            }

            switch (_pos)
            {
                case Position.TopLeft:
                case Position.Left:
                case Position.BottomLeft:
                    _align = Alignment.Left;
                    break;
                case Position.Top:
                case Position.Centre:
                case Position.Bottom:
                    _align = Alignment.Centre;
                    break;
                case Position.TopRight:
                case Position.Right:
                case Position.BottomRight:
                    _align = Alignment.Right;
                    break;
            }
        }

        private void SetTextAlignment(MsoTextEffectAlignment alignment)
        {
            foreach (Shape shape in TextShapes)
            {
                shape.TextEffect.Alignment = alignment;
            }
        }

        private void SetupTextBoxesPosition(TextBoxInfo boxesInfo)
        {
            switch (_pos)
            {
                case Position.TopLeft:
                case Position.Left:
                case Position.BottomLeft:
                    _left = Margin;
                    break;
                case Position.Top:
                case Position.Centre:
                case Position.Bottom:
                    _left = _slideWidth/2 - boxesInfo.Width/2;
                    break;
                case Position.TopRight:
                case Position.Right:
                case Position.BottomRight:
                    _left = _slideWidth - boxesInfo.Width - Margin;
                    break;
            }
            switch (_pos)
            {
                case Position.TopLeft:
                case Position.Top:
                case Position.TopRight:
                    _top = Margin;
                    break;
                case Position.Left:
                case Position.Centre:
                case Position.Right:
                    _top = _slideHeight/2 - boxesInfo.Height/2;
                    break;
                case Position.BottomLeft:
                case Position.Bottom:
                case Position.BottomRight:
                    _top = _slideHeight - boxesInfo.Height - Margin;
                    break;
            }
        }

        private TextBoxInfo GetTextBoxInfo(Shape textShape)
        {
            TextBoxInfo result = new TextBoxInfo();
            TextRange2 paragraphs = textShape.TextFrame2.TextRange.Paragraphs;
            float rightMost = 0f;
            float bottomMost = 0f;
            foreach (TextRange2 textRange in paragraphs)
            {
                TextRange2 paragraph = textRange.TrimText();
                if (StringUtil.IsNotEmpty(paragraph.Text))
                {
                    result.Left = paragraph.BoundLeft < result.Left ? paragraph.BoundLeft : result.Left;
                    result.Top = paragraph.BoundTop < result.Top ? paragraph.BoundTop : result.Top;
                    rightMost = paragraph.BoundLeft + paragraph.BoundWidth > rightMost
                        ? paragraph.BoundLeft + paragraph.BoundWidth
                        : rightMost;
                    bottomMost = paragraph.BoundTop + paragraph.BoundHeight > bottomMost
                        ? paragraph.BoundTop + paragraph.BoundHeight
                        : bottomMost;
                }
            }
            result.Width = rightMost - result.Left;
            result.Height = bottomMost - result.Top;
            AddMargin(result, ExtendMargin);
            return result;
        }

        private TextBoxInfo GetTextBoxesInfo(IEnumerable<Shape> textShapes)
        {
            TextBoxInfo result = new TextBoxInfo();
            float rightMost = 0f;
            float bottomMost = 0f;
            foreach (TextBoxInfo partialResult in textShapes.Select(GetTextBoxInfo))
            {
                result.Left = partialResult.Left < result.Left ? partialResult.Left : result.Left;
                result.Top = partialResult.Top < result.Top ? partialResult.Top : result.Top;
                rightMost = partialResult.Left + partialResult.Width > rightMost
                        ? partialResult.Left + partialResult.Width
                        : rightMost;
                bottomMost = partialResult.Top + partialResult.Height > bottomMost
                    ? partialResult.Top + partialResult.Height
                    : bottomMost;
            }
            result.Width = rightMost - result.Left;
            result.Height = bottomMost - result.Top;
            return result;
        }

        // rule:
        // top > left > name: Title > name: Subtitle > name: Text > other
        // 
        // exp sorted result:
        // most top textbox at the first element,
        // most bottom textbox at the last element
        private void SortTextBoxes()
        {
            TextShapes.Sort((shape1, shape2) =>
            {
                if ((int)(shape2.Top - shape1.Top) != 0)
                {
                    return (int) (shape1.Top - shape2.Top);
                }
                if ((int)(shape2.Left - shape1.Left) != 0)
                {
                    return (int) (shape1.Left - shape2.Left);
                }
                if (shape1.Name.StartsWith("Title"))
                {
                    return -1;
                }
                if (shape2.Name.StartsWith("Title"))
                {
                    return 1;
                }
                if (shape1.Name.StartsWith("Subtitle"))
                {
                    return -1;
                }
                if (shape2.Name.StartsWith("Subtitle"))
                {
                    return 1;
                }
                if (shape1.Name.StartsWith("Text"))
                {
                    return -1;
                }
                if (shape2.Name.StartsWith("Text"))
                {
                    return 1;
                }
                return -1;
            });
        }

        # endregion
    }
}
