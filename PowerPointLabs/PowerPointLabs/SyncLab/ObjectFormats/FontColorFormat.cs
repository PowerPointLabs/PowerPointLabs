﻿using System;
using System.Drawing;
using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FontColorFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return formatShape.HasTextFrame == MsoTriState.msoTrue;
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Font Color");
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
            
            int guessedColor = ShapeUtil.GuessTextColor(formatShape);
            shape.Fill.ForeColor.RGB = guessedColor;
            shape.Fill.BackColor.RGB = guessedColor;
            shape.Fill.Solid();
            Bitmap image = GraphicsUtil.ShapeToBitmap(shape);
            shape.Delete();
            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                int guessedColor = ShapeUtil.GuessTextColor(formatShape);
                newShape.TextFrame.TextRange.Font.Color.RGB = guessedColor;
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
    }
}
