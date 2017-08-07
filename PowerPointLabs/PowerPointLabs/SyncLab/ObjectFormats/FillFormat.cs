using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FillFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            var duplicateShape = formatShape.Duplicate()[1];
            bool canCopy = Sync(formatShape, duplicateShape);
            duplicateShape.Delete();
            return canCopy;
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Fill");
            }
        }

        public static Bitmap DisplayImage(Shape formatShape)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Shape shape = shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 
                    SyncFormatConstants.DisplayImageSize.Width,
                    SyncFormatConstants.DisplayImageSize.Height);
            shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            SyncFormat(formatShape, shape);
            Bitmap image = new Bitmap(GraphicsUtil.ShapeToBitmap(shape));
            shape.Delete();
            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                if (formatShape.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillPatterned)
                {
                    newShape.Fill.Patterned(formatShape.Fill.Pattern);
                    newShape.Fill.ForeColor.RGB = formatShape.Fill.ForeColor.RGB;
                    newShape.Fill.BackColor.RGB = formatShape.Fill.BackColor.RGB;
                }
                else if (formatShape.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillBackground)
                {
                    newShape.Fill.Background();
                }
                else if (formatShape.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillSolid)
                {
                    newShape.Fill.Solid();
                    newShape.Fill.ForeColor.RGB = formatShape.Fill.ForeColor.RGB;
                }
                else if (formatShape.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillGradient)
                {
                    SyncGradient(formatShape, newShape);
                }
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }

        private static void SyncGradient(Shape formatShape, Shape newShape) //should return bool?
        {
            var gradientColorType = formatShape.Fill.GradientColorType;

            if (gradientColorType == Microsoft.Office.Core.MsoGradientColorType.msoGradientOneColor)
            {
                newShape.Fill.OneColorGradient(formatShape.Fill.GradientStyle, formatShape.Fill.GradientVariant, formatShape.Fill.GradientDegree);
                SyncInitialGradientStops(formatShape, newShape);
            }
            else if (gradientColorType == Microsoft.Office.Core.MsoGradientColorType.msoGradientTwoColors)
            {
                newShape.Fill.TwoColorGradient(formatShape.Fill.GradientStyle, formatShape.Fill.GradientVariant);
                SyncInitialGradientStops(formatShape, newShape);
            }
            else if (gradientColorType == Microsoft.Office.Core.MsoGradientColorType.msoGradientPresetColors)
            {
                newShape.Fill.PresetGradient(formatShape.Fill.GradientStyle, formatShape.Fill.GradientVariant, formatShape.Fill.PresetGradientType);
            }
            else if (gradientColorType == Microsoft.Office.Core.MsoGradientColorType.msoGradientMultiColor)
            {
                int formatGradientVarient = 1;
                try
                {
                    formatGradientVarient = formatShape.Fill.GradientVariant;
                }
                catch
                {
                    formatGradientVarient = 1;
                }
                newShape.Fill.OneColorGradient(formatShape.Fill.GradientStyle, formatGradientVarient, 0);
                SyncInitialGradientStops(formatShape, newShape);
                for (int i = 3; i <= formatShape.Fill.GradientStops.Count; i++)
                {
                    newShape.Fill.GradientStops.Insert(formatShape.Fill.GradientStops[i].Color.RGB,
                        formatShape.Fill.GradientStops[i].Position, formatShape.Fill.GradientStops[i].Transparency);
                }

                try
                {
                    newShape.Fill.GradientAngle = formatShape.Fill.GradientAngle;
                }
                catch (Exception)
                {
                    //gradient has no angle
                }
            }
        }

        private static void SyncInitialGradientStops(Shape formatShape, Shape newShape)
        {
            newShape.Fill.GradientStops[1].Color.RGB = formatShape.Fill.GradientStops[1].Color.RGB;
            newShape.Fill.GradientStops[1].Position = formatShape.Fill.GradientStops[1].Position;

            newShape.Fill.GradientStops[2].Color.RGB = formatShape.Fill.GradientStops[2].Color.RGB;
            newShape.Fill.GradientStops[2].Position = formatShape.Fill.GradientStops[2].Position;

            newShape.Fill.RotateWithObject = formatShape.Fill.RotateWithObject;
        }
    }
}
