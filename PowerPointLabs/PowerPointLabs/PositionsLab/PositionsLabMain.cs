using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.PositionsLab
{
    class PositionsLabMain
    {
        #region API
        public static void SnapVertical()
        {
            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide as PowerPointSlide;
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;

            for (int i = 1; i <= selectedShapes.Count; i++)
            {
                SnapShapeVertical(selectedShapes[i]);
            }
        }

        public static void SnapHorizontal()
        {
            var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide as PowerPointSlide;
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;

            for (int i = 1; i <= selectedShapes.Count; i++)
            {
                SnapShapeHorizontal(selectedShapes[i]);
            }
        }

        public static void SnapShapeVertical(Shape shape)
        {
            if (IsVertical(shape))
            {
                SnapTo0Or180(shape);
            }
            else
            {
                SnapTo90Or270(shape);
            }
        }

        public static void SnapShapeHorizontal(Shape shape)
        {
            if (IsVertical(shape))
            {
                SnapTo90Or270(shape);
            }
            else
            {
                SnapTo0Or180(shape);
            }
        }

        private static void SnapTo0Or180 (Shape shape)
        {
            float rotation = shape.Rotation;

            if (rotation >= 90 && rotation < 270)
            {
                shape.Rotation = 180;
            }
            else
            {
                shape.Rotation = 0;
            }
        }

        private static void SnapTo90Or270(Shape shape)
        {
            float rotation = shape.Rotation;

            if (rotation >= 0 && rotation < 180)
            {
                shape.Rotation = 90;
            }
            else
            {
                shape.Rotation = 270;
            }
        }

        private static bool IsVertical(Shape shape)
        {
            return shape.Height > shape.Width;
        }
        #endregion
    }
}
