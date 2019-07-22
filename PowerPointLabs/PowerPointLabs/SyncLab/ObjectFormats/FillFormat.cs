using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FillFormat: Format 
    {
        public override bool CanCopy(Shape formatShape)
        {
            Shape duplicateShape = formatShape.Duplicate()[1];
            bool canCopy = Sync(formatShape, duplicateShape);
            duplicateShape.SafeDelete();
            return canCopy;
        }

        public override void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Fill");
            }
        }

        public override Bitmap DisplayImage(Shape formatShape)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Shape shape = shapes.AddShape(
                    Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, 0, 0, 
                    SyncFormatConstants.DisplayImageSize.Width,
                    SyncFormatConstants.DisplayImageSize.Height);
            shape.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
            SyncFormat(formatShape, shape);
            Bitmap image = new Bitmap(GraphicsUtil.ShapeToBitmap(shape));
            shape.SafeDelete();
            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                // force msoFillMixed to msoFillSolid
                // freshly created textboxes have the msoFillMixed type 
                // otherwise, msoFillMixed only appears when multiple shapes are selected
                // manual conversion is needed as msoFillMixed textboxes risk system forced conversions to msoFillSolid
                // system forced conversions will set fill color to black
                //
                // lines also have the msoFillMixed type
                // they have no fill, throwing an exception in the following if block
                // this is desired behavior, disabling FillFormat for lines
                if (formatShape.Fill.Type == Microsoft.Office.Core.MsoFillType.msoFillMixed)
                {
                    int oldColor = formatShape.Fill.ForeColor.RGB;
                    float oldTransparency = formatShape.Fill.Transparency;
                    formatShape.Fill.Solid();
                    formatShape.Fill.ForeColor.RGB = oldColor;
                    formatShape.Fill.Transparency = oldTransparency;
                }
                
                switch (formatShape.Fill.Type)
                {
                    case Microsoft.Office.Core.MsoFillType.msoFillPatterned:
                        newShape.Fill.Patterned(formatShape.Fill.Pattern);
                        newShape.Fill.ForeColor.RGB = formatShape.Fill.ForeColor.RGB;
                        newShape.Fill.BackColor.RGB = formatShape.Fill.BackColor.RGB;
                        break;
                    case Microsoft.Office.Core.MsoFillType.msoFillBackground:
                        newShape.Fill.Background();
                        break;
                    case Microsoft.Office.Core.MsoFillType.msoFillSolid:
                        newShape.Fill.Solid();
                        newShape.Fill.ForeColor.RGB = formatShape.Fill.ForeColor.RGB;
                        break;
                    case Microsoft.Office.Core.MsoFillType.msoFillGradient:
                        SyncGradient(formatShape, newShape);
                        break;
                    default:
                        break;
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Sync FillFormat");
                return false;
            }
            return true;
        }

        private static void SyncGradient(Shape formatShape, Shape newShape) //should return bool?
        {
            Microsoft.Office.Core.MsoGradientColorType gradientColorType = formatShape.Fill.GradientColorType;

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
