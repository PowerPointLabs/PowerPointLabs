using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    public static class PowerpointExtension
    {
        // groups the shaperange only if there is 2 or more elements
        public static Shape SafeGroup(this ShapeRange range)
        {
            if (range.Count > 1)
            {
                return range.Group();
            }
            else if (range.Count == 1)
            {
                return range[1];
            }
            else
            {
                return null;
            }
        }

        // copies placeholder textboxes safely
        public static Shape SafeCopy(this Shapes shapes, Shape shape)
        {
            if (shape.IsEmptyPlaceholder())
            {
                Shape newShape = shapes.AddTextbox(shape.TextFrame.Orientation, shape.Left, shape.Top, shape.Width, shape.Height);
                newShape.CopyColorFrom(shape);
                return newShape;
            }
            shape.Copy();
            return shapes.Paste()[1];
        }

        // Referred to code in ColorsLabPaneWPFxaml
        public static void CopyColorFrom(this Shape newShape, Shape shape)
        {
            // TODO: Brute force copy fill and line
            // Some stuff not safe
            newShape.Fill.ForeColor = shape.Fill.ForeColor;

            newShape.Line.ForeColor = shape.Line.ForeColor;
            newShape.Line.DashStyle = shape.Line.DashStyle;
            newShape.Line.Style = shape.Line.Style;
            newShape.Line.Transparency = shape.Line.Transparency;
            newShape.Line.Visible = shape.Line.Visible;
            newShape.Line.Weight = shape.Line.Weight;
            newShape.Line.BackColor = shape.Line.BackColor;
            newShape.Line.Style = shape.Line.Style;

            if (shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                Font newShapeFont = newShape.TextFrame.TextRange.Font;
                Font shapeFont = shape.TextFrame.TextRange.Font;
                newShapeFont.CopyFontFrom(shapeFont);
            }
        }

        // a brute force copy
        public static void CopyFontFrom(this Font font, Font other)
        {
            font.AutoRotateNumbers = other.AutoRotateNumbers;
            font.BaselineOffset = other.BaselineOffset;
            font.Bold = other.Bold;
            font.Emboss = other.Emboss;
            font.Italic = other.Italic;
            font.Shadow = other.Shadow;
            font.Subscript = other.Subscript;
            font.Superscript = other.Superscript;
            font.Underline = other.Underline;
            font.Color.CopyColorFrom(other.Color);
        }

        public static void CopyColorFrom(this ColorFormat color, ColorFormat other)
        {
            color.Brightness = other.Brightness;
            color.RGB = other.RGB;
            color.SchemeColor = other.SchemeColor;
            color.TintAndShade = other.TintAndShade;
        }

        public static bool IsEmptyPlaceholder(this Shape shape)
        {
            return shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder &&
                shape.TextFrame.TextRange.Text.Length == 0;
        }
    }
}
