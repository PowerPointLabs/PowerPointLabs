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
using System.Diagnostics;
using Drawing = System.Drawing;

namespace PowerPointLabs.PositionsLab
{
    class PositionsLabMain
    {
        #region API

        #region Align
        public static void AlignLeft()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;

            if (selectedShapes.Count < 2)
            {
                //Error
                return;
            }

            Shape refShape = selectedShapes[1];
            Drawing.PointF[] allPointsOfRef = GetRealCoordinates(refShape);
            Drawing.PointF leftMostRef = LeftMostPoint(allPointsOfRef);
            
            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF[] allPoints = GetRealCoordinates(s);
                Drawing.PointF leftMost = LeftMostPoint(allPoints);
                s.IncrementLeft(leftMostRef.X - leftMost.X);
            }
        }

        public static void AlignRight()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;

            if (selectedShapes.Count < 2)
            {
                //Error
                return;
            }

            Shape refShape = selectedShapes[1];
            Drawing.PointF[] allPointsOfRef = GetRealCoordinates(refShape);
            Drawing.PointF rightMostRef = RightMostPoint(allPointsOfRef);

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF[] allPoints = GetRealCoordinates(s);
                Drawing.PointF rightMost = RightMostPoint(allPoints);
                s.IncrementLeft(rightMostRef.X - rightMost.X);
            }
        }

        public static void AlignTop()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;

            if (selectedShapes.Count < 2)
            {
                //Error
                return;
            }

            Shape refShape = selectedShapes[1];
            Drawing.PointF[] allPointsOfRef = GetRealCoordinates(refShape);
            Drawing.PointF topMostRef = TopMostPoint(allPointsOfRef);

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF[] allPoints = GetRealCoordinates(s);
                Drawing.PointF topMost = TopMostPoint(allPoints);
                s.IncrementTop(topMostRef.Y - topMost.Y);
            }
        }

        public static void AlignBottom()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;

            if (selectedShapes.Count < 2)
            {
                //Error
                return;
            }

            Shape refShape = selectedShapes[1];
            Drawing.PointF[] allPointsOfRef = GetRealCoordinates(refShape);
            Drawing.PointF lowestRef = LowestPoint(allPointsOfRef);

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF[] allPoints = GetRealCoordinates(s);
                Drawing.PointF lowest = LowestPoint(allPoints);
                s.IncrementTop(lowestRef.Y - lowest.Y);
            }
        }

        public static void AlignMiddle()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;

            if (selectedShapes.Count < 2)
            {
                //Error
                return;
            }

            Shape refShape = selectedShapes[1];
            Drawing.PointF originRef = GetOrigin(refShape);

            for (int i=2; i<= selectedShapes.Count; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF origin = GetOrigin(s);
                s.IncrementTop(originRef.Y - origin.Y);
            }
        }

        public static void AlignCenter()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;

            if (selectedShapes.Count < 2)
            {
                //Error
                return;
            }

            Shape refShape = selectedShapes[1];
            Drawing.PointF originRef = GetOrigin(refShape);

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF origin = GetOrigin(s);
                s.IncrementLeft(originRef.X - origin.X);
                s.IncrementTop(originRef.Y - origin.Y);
            }
        }
        #endregion


        public static void SnapVertical()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;

            for (int i = 1; i <= selectedShapes.Count; i++)
            {
                SnapShapeVertical(selectedShapes[i]);
            }
        }

        public static void SnapHorizontal()
        {
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

        #region Util
        private static Drawing.PointF[] GetRealCoordinates(Shape s)
        {
            float rotation = s.Rotation;

            Drawing.PointF s1 = new Drawing.PointF(s.Left, s.Top);
            Drawing.PointF s2 = new Drawing.PointF(s.Left + s.Width, s.Top);
            Drawing.PointF s3 = new Drawing.PointF(s.Left + s.Width, s.Top + s.Height);
            Drawing.PointF s4 = new Drawing.PointF(s.Left, s.Top + s.Height);
            Drawing.PointF origin = GetOrigin(s);

            Drawing.PointF rotated1 = RotatePoint(s1, origin, rotation);
            Drawing.PointF rotated2 = RotatePoint(s2, origin, rotation);
            Drawing.PointF rotated3 = RotatePoint(s3, origin, rotation);
            Drawing.PointF rotated4 = RotatePoint(s4, origin, rotation);

            return new Drawing.PointF[] { rotated1, rotated2, rotated3, rotated4 };

        }

        private static Drawing.PointF GetOrigin(Shape s)
        {
            return new Drawing.PointF(s.Left + s.Width / 2, s.Top + s.Height / 2);
        }

        private static Drawing.PointF RotatePoint(Drawing.PointF p, Drawing.PointF origin, float rotation)
        {
            double rotationInRadian = DegreeToRadian(rotation);
            double rotatedX = Math.Cos(rotationInRadian) * (p.X - origin.X) - Math.Sin(rotationInRadian) * (p.Y - origin.Y) + origin.X;
            double rotatedY = Math.Sin(rotationInRadian) * (p.X - origin.X) - Math.Cos(rotationInRadian) * (p.Y - origin.Y) + origin.Y;

            return new Drawing.PointF((float)rotatedX, (float)rotatedY);
        }

        private static double DegreeToRadian(float angle)
        {
            return angle / 180.0 * Math.PI;
        }

        private static Drawing.PointF LeftMostPoint(Drawing.PointF[] coords)
        {
            Drawing.PointF leftMost = new Drawing.PointF();

            for (int i = 0; i < coords.Length; i++)
            {
                if (leftMost.IsEmpty)
                {
                    leftMost = coords[i];
                }
                else
                {
                    if (coords[i].X < leftMost.X)
                    {
                        leftMost = coords[i];
                    }
                }
            }

            return leftMost;
        }

        private static Drawing.PointF RightMostPoint(Drawing.PointF[] coords)
        {
            Drawing.PointF rightMost = new Drawing.PointF();

            for (int i = 0; i < coords.Length; i++)
            {
                if (rightMost.IsEmpty)
                {
                    rightMost = coords[i];
                }
                else
                {
                    if (coords[i].X > rightMost.X)
                    {
                        rightMost = coords[i];
                    }
                }
            }

            return rightMost;
        }

        private static Drawing.PointF TopMostPoint(Drawing.PointF[] coords)
        {
            Drawing.PointF topMost = new Drawing.PointF();

            for (int i = 0; i < coords.Length; i++)
            {
                if (topMost.IsEmpty)
                {
                    topMost = coords[i];
                }
                else
                {
                    if (coords[i].Y < topMost.Y)
                    {
                        topMost = coords[i];
                    }
                }
            }

            return topMost;
        }

        private static Drawing.PointF LowestPoint(Drawing.PointF[] coords)
        {
            Drawing.PointF lowest = new Drawing.PointF();

            for (int i = 0; i < coords.Length; i++)
            {
                if (lowest.IsEmpty)
                {
                    lowest = coords[i];
                }
                else
                {
                    if (coords[i].Y > lowest.Y)
                    {
                        lowest = coords[i];
                    }
                }
            }

            return lowest;
        }

        private static double GetUnrotatedLeftGivenRotatedLeft(Shape s, float rotatedLeft)
        {
            double rotationInRadian = DegreeToRadian(s.Rotation);
            return rotatedLeft + Math.Cos(rotationInRadian) * (s.Width / 2) - Math.Sin(rotationInRadian) * (s.Height / 2) - s.Width / 2;
        }
        #endregion
    }
}
