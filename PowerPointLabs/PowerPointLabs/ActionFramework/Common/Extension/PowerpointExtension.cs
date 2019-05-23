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
            newShape.Fill.ForeColor = shape.Fill.ForeColor;
            newShape.Line.ForeColor = shape.Line.ForeColor;
            if (shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                newShape.TextFrame.TextRange.TrimText().Font.Color.RGB = shape.TextFrame.TextRange.TrimText().Font.Color.RGB;
            }
        }

        public static bool IsEmptyPlaceholder(this Shape shape)
        {
            return shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder &&
                shape.TextFrame.TextRange.Text.Length == 0;
        }
    }
}
