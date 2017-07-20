﻿using System;
using System.Drawing;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Utils;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class LineDashTypeFormat
    {
        public static bool CanCopy(Shape formatShape)
        {
            return Sync(formatShape, formatShape);
        }

        public static void SyncFormat(Shape formatShape, Shape newShape)
        {
            if (!Sync(formatShape, newShape))
            {
                Logger.Log(newShape.Type + " unable to sync Dash Style");
            }
        }

        public static Bitmap DisplayImage(Shape formatShape)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Shape shape = shapes.AddLine(
                0, SyncFormatConstants.DisplayImageSize.Height,
                SyncFormatConstants.DisplayImageSize.Width, 0);
            SyncFormat(formatShape, shape);
            shape.Line.ForeColor.RGB = SyncFormatConstants.ColorBlack;
            shape.Line.Weight = SyncFormatConstants.DisplayLineWeight;
            Bitmap image = GraphicsUtil.ShapeToBitmap(shape);
            shape.Delete();
            return image;
        }

        private static bool Sync(Shape formatShape, Shape newShape)
        {
            try
            {
                newShape.Line.DashStyle = formatShape.Line.DashStyle;
            }
            catch (Exception)
            {
                return false;
            }
            return true;
        }
    }
}
