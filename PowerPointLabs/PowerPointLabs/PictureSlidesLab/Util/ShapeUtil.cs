using System;

using Microsoft.Office.Core;

using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PictureSlidesLab.Util
{
    public class ShapeUtil
    {
        public static void ChangeName(PowerPoint.Shape shape, EffectName effectName, string shapeNamePrefix)
        {
            shape.Name = shapeNamePrefix + "_" + effectName + "_" + Guid.NewGuid().ToString().Substring(0, 7);
        }

        public static void AddTag(PowerPoint.Shape shape, string tagName, String value)
        {
            if (StringUtil.IsEmpty(shape.Tags[tagName]) && value != null)
            {
                shape.Tags.Add(tagName, value);
            }
        }

        public static PowerPoint.Shape GetTextShapeToProcess(PowerPoint.ShapeRange shapes)
        {
            foreach (PowerPoint.Shape shape in shapes)
            {
                if (shape.Type != MsoShapeType.msoPlaceholder
                    || shape.TextFrame.HasText == MsoTriState.msoFalse)
                {
                    continue;
                }

                switch (shape.PlaceholderFormat.Type)
                {
                    case PowerPoint.PpPlaceholderType.ppPlaceholderTitle:
                    case PowerPoint.PpPlaceholderType.ppPlaceholderCenterTitle:
                    case PowerPoint.PpPlaceholderType.ppPlaceholderVerticalTitle:
                        return shape;
                }
            }

            foreach (PowerPoint.Shape shape in shapes)
            {
                if ((shape.Type != MsoShapeType.msoPlaceholder
                        && shape.Type != MsoShapeType.msoTextBox)
                        || shape.TextFrame.HasText == MsoTriState.msoFalse
                        || StringUtil.IsNotEmpty(shape.Tags[Tag.ImageReference]))
                {
                    continue;
                }

                if (shape.Type == MsoShapeType.msoPlaceholder
                    && (shape.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderSlideNumber
                        || shape.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderFooter
                        || shape.PlaceholderFormat.Type == PowerPoint.PpPlaceholderType.ppPlaceholderHeader))
                {
                    continue;
                }

                return shape;
            }
            return null;
        }

        public static PowerPoint.Shape GetTextShapeToProcess(PowerPoint.Shapes shapes)
        {
            return GetTextShapeToProcess(shapes.Range());
        }
    }
}
