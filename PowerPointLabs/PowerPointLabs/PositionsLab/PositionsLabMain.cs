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

        private static bool _useSlideAsReference = false;

        #region API

        #region Class Methods

        /// <summary>
        /// Tells the Positions Lab to use the slide as the reference point for the methods
        /// </summary>
        public static void ReferToSlide()
        {
            _useSlideAsReference = true;
        }

        /// <summary>
        /// Tells the Positions Lab to use reference shapes for the methods
        /// </summary>
        public static void ReferToShape()
        {
            _useSlideAsReference = false;
        }
        #endregion

        #region Align
        public static void AlignLeft()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;

            if (_useSlideAsReference)
            {
                foreach (Shape s in selectedShapes)
                {
                    Drawing.PointF[] allPointsOfShape = Graphics.GetRealCoordinates(s);
                    Drawing.PointF leftMost = Graphics.LeftMostPoint(allPointsOfShape);
                    s.IncrementLeft(-leftMost.X);
                }
            }
            else
            {
                if (selectedShapes.Count < 2)
                {
                    //Error
                    return;
                }

                Shape refShape = selectedShapes[1];
                Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);
                Drawing.PointF leftMostRef = Graphics.LeftMostPoint(allPointsOfRef);

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    Shape s = selectedShapes[i];
                    Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                    Drawing.PointF leftMost = Graphics.LeftMostPoint(allPoints);
                    s.IncrementLeft(leftMostRef.X - leftMost.X);
                }
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
            Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);
            Drawing.PointF rightMostRef = Graphics.RightMostPoint(allPointsOfRef);

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                Drawing.PointF rightMost = Graphics.RightMostPoint(allPoints);
                s.IncrementLeft(rightMostRef.X - rightMost.X);
            }
        }

        public static void AlignTop()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;

            if (_useSlideAsReference)
            {
                foreach (Shape s in selectedShapes)
                {
                    Drawing.PointF[] allPointsOfShape = Graphics.GetRealCoordinates(s);
                    Drawing.PointF topMost = Graphics.TopMostPoint(allPointsOfShape);
                    s.IncrementTop(-topMost.Y);
                }
            }
            else
            {
                if (selectedShapes.Count < 2)
                {
                    //Error
                    return;
                }

                Shape refShape = selectedShapes[1];
                Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);
                Drawing.PointF topMostRef = Graphics.TopMostPoint(allPointsOfRef);

                for (int i = 2; i <= selectedShapes.Count; i++)
                {
                    Shape s = selectedShapes[i];
                    Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                    Drawing.PointF topMost = Graphics.TopMostPoint(allPoints);
                    s.IncrementTop(topMostRef.Y - topMost.Y);
                }
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
            Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);
            Drawing.PointF lowestRef = Graphics.BottomMostPoint(allPointsOfRef);

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                Drawing.PointF lowest = Graphics.BottomMostPoint(allPoints);
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
            Drawing.PointF originRef = Graphics.GetCenterPoint(refShape);

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF origin = Graphics.GetCenterPoint(s);
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
            Drawing.PointF originRef = Graphics.GetCenterPoint(refShape);

            for (int i = 2; i <= selectedShapes.Count; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF origin = Graphics.GetCenterPoint(s);
                s.IncrementLeft(originRef.X - origin.X);
                s.IncrementTop(originRef.Y - origin.Y);
            }
        }

        #endregion

        #region Snap
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

        #region Adjoin
        public static void AdjoinHorizontal()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;

            if (selectedShapes.Count < 2)
            {
                //Error
                return;
            }

            Shape refShape = selectedShapes[1];
            List<Shape> sortedShapes = Graphics.SortShapesByLeft(selectedShapes);
            int refShapeIndex = sortedShapes.IndexOf(refShape);

            Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);
            Drawing.PointF centerOfRef = Graphics.GetCenterPoint(refShape);

            float mostLeft = Graphics.LeftMostPoint(allPointsOfRef).X;
            //For all shapes left of refShape, adjoin them from closest to refShape
            for (int i = refShapeIndex - 1; i >= 0; i--)
            {
                Shape neighbour = sortedShapes[i];
                Drawing.PointF[] allPointsOfNeighbour = Graphics.GetRealCoordinates(neighbour);
                float rightOfShape = Graphics.RightMostPoint(allPointsOfNeighbour).X;
                neighbour.IncrementLeft(mostLeft - rightOfShape);
                neighbour.IncrementTop(centerOfRef.Y - Graphics.GetCenterPoint(neighbour).Y);

                mostLeft = Graphics.LeftMostPoint(allPointsOfNeighbour).X + mostLeft - rightOfShape;
            }

            float mostRight = Graphics.RightMostPoint(allPointsOfRef).X;
            //For all shapes right of refShape, adjoin them from closest to refShape
            for (int i = refShapeIndex + 1; i < sortedShapes.Count; i++)
            {
                Shape neighbour = sortedShapes[i];
                Drawing.PointF[] allPointsOfNeighbour = Graphics.GetRealCoordinates(neighbour);
                float leftOfShape = Graphics.LeftMostPoint(allPointsOfNeighbour).X;
                neighbour.IncrementLeft(mostRight - leftOfShape);
                neighbour.IncrementTop(centerOfRef.Y - Graphics.GetCenterPoint(neighbour).Y);

                mostRight = Graphics.RightMostPoint(allPointsOfNeighbour).X + mostRight - leftOfShape;
            }
        }

        public static void AdjoinVertical()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;

            if (selectedShapes.Count < 2)
            {
                //Error
                return;
            }

            Shape refShape = selectedShapes[1];
            List<Shape> sortedShapes = Graphics.SortShapesByTop(selectedShapes);
            int refShapeIndex = sortedShapes.IndexOf(refShape);

            Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);
            Drawing.PointF centerOfRef = Graphics.GetCenterPoint(refShape);

            float mostTop = Graphics.TopMostPoint(allPointsOfRef).Y;
            //For all shapes above refShape, adjoin them from closest to refShape
            for (int i = refShapeIndex - 1; i >= 0; i--)
            {
                Shape neighbour = sortedShapes[i];
                Drawing.PointF[] allPointsOfNeighbour = Graphics.GetRealCoordinates(neighbour);
                float bottomOfShape = Graphics.BottomMostPoint(allPointsOfNeighbour).Y;
                neighbour.IncrementLeft(centerOfRef.X - Graphics.GetCenterPoint(neighbour).X);
                neighbour.IncrementTop(mostTop - bottomOfShape);

                mostTop = Graphics.TopMostPoint(allPointsOfNeighbour).Y + mostTop - bottomOfShape;
            }

            float lowest = Graphics.BottomMostPoint(allPointsOfRef).Y;
            //For all shapes right of refShape, adjoin them from closest to refShape
            for (int i = refShapeIndex + 1; i < sortedShapes.Count; i++)
            {
                Shape neighbour = sortedShapes[i];
                Drawing.PointF[] allPointsOfNeighbour = Graphics.GetRealCoordinates(neighbour);
                float topOfShape = Graphics.TopMostPoint(allPointsOfNeighbour).Y;
                neighbour.IncrementLeft(centerOfRef.X - Graphics.GetCenterPoint(neighbour).X);
                neighbour.IncrementTop(lowest - topOfShape);

                lowest = Graphics.BottomMostPoint(allPointsOfNeighbour).Y + lowest - topOfShape;
            }
        }
        #endregion

        #region Swap
        public static void Swap()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;

            if (selectedShapes.Count < 2)
            {
                //Error
                return;
            }

            List<Shape> sortedShapes = Graphics.SortShapesByLeft(selectedShapes);
            Drawing.PointF firstPos = Graphics.GetCenterPoint(sortedShapes[0]);

            for (int i = 0; i < sortedShapes.Count; i++)
            {
                Shape currentShape = sortedShapes[i];
                if (i < sortedShapes.Count - 1)
                {
                    Drawing.PointF currentPos = Graphics.GetCenterPoint(currentShape);
                    Drawing.PointF nextPos = Graphics.GetCenterPoint(sortedShapes[i + 1]);

                    currentShape.IncrementLeft(nextPos.X - currentPos.X);
                    currentShape.IncrementTop(nextPos.Y - currentPos.Y);
                }
                else
                {
                    Drawing.PointF currentPos = Graphics.GetCenterPoint(currentShape);
                    currentShape.IncrementLeft(firstPos.X - currentPos.X);
                    currentShape.IncrementTop(firstPos.Y - currentPos.Y);
                }
            }
        }
        #endregion

        #region Distribute
        public static void DistributeHorizontal()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;
            var shapeCount = selectedShapes.Count;
            if (shapeCount < 2)
            {
                //Error
                return;
            }

            Shape refShape = selectedShapes[1];
            Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);
            Drawing.PointF leftMostRef = Graphics.LeftMostPoint(allPointsOfRef);
            Drawing.PointF rightMostRef = Graphics.RightMostPoint(allPointsOfRef);

            var horizontalDistanceInRef = Graphics.RealWidth(allPointsOfRef);
            var spaceBetweenShapes = horizontalDistanceInRef;

            for (int i = 2; i <= shapeCount; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                var shapeWidth = Graphics.RealWidth(allPoints);
                spaceBetweenShapes -= shapeWidth;
            }

            // TODO: guard against spaceBetweenShapes < 0

            spaceBetweenShapes /= shapeCount;

            for (int i = 2; i <= shapeCount; i++)
            {
                Shape currShape = selectedShapes[i];
                Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(currShape);
                Drawing.PointF leftMost = Graphics.LeftMostPoint(allPoints);
                refShape = selectedShapes[i - 1];
                allPointsOfRef = Graphics.GetRealCoordinates(refShape);

                if (i == 2)
                {
                    leftMostRef = Graphics.LeftMostPoint(allPointsOfRef);
                    currShape.IncrementLeft(leftMostRef.X - leftMost.X + spaceBetweenShapes);
                }
                else
                {
                    rightMostRef = Graphics.RightMostPoint(allPointsOfRef);
                    currShape.IncrementLeft(rightMostRef.X - leftMost.X + spaceBetweenShapes);
                }
            }
        } 

        public static void DistributeVertical()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;
            var shapeCount = selectedShapes.Count;
            if (shapeCount < 2)
            {
                //Error
                return;
            }

            Shape refShape = selectedShapes[1];
            Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);
            Drawing.PointF topMostRef = Graphics.TopMostPoint(allPointsOfRef);
            Drawing.PointF lowestRef = Graphics.BottomMostPoint(allPointsOfRef);

            var verticalDistanceInRef = Graphics.RealHeight(allPointsOfRef);
            var spaceBetweenShapes = verticalDistanceInRef;

            for (int i = 2; i <= shapeCount; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                var shapeHeight = Graphics.RealHeight(allPoints);
                spaceBetweenShapes -= shapeHeight;
            }

            // TODO: guard against spaceBetweenShapes < 0

            spaceBetweenShapes /= shapeCount;

            for (int i = 2; i <= shapeCount; i++)
            {
                Shape currShape = selectedShapes[i];
                Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(currShape);
                Drawing.PointF topMost = Graphics.TopMostPoint(allPoints);
                refShape = selectedShapes[i - 1];
                allPointsOfRef = Graphics.GetRealCoordinates(refShape);

                if (i == 2)
                {
                    topMostRef = Graphics.TopMostPoint(allPointsOfRef);
                    currShape.IncrementTop(topMostRef.Y - topMost.Y + spaceBetweenShapes);
                }
                else
                {
                    lowestRef = Graphics.BottomMostPoint(allPointsOfRef);
                    currShape.IncrementTop(lowestRef.Y - topMost.Y + spaceBetweenShapes);
                }
            }
        }

        public static void DistributeCenter()
        {
            DistributeHorizontal();
            DistributeVertical();
        }

        public static void DistributeShapes()
        {
            var selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange as PowerPoint.ShapeRange;
            var shapeCount = selectedShapes.Count;

            if (shapeCount < 2)
            {
                //Error
                return;
            }

            if (shapeCount == 2)
            {
                return;
            }

            Shape firstRef = selectedShapes[1];
            Shape lastRef = selectedShapes[selectedShapes.Count];
            Shape refShape = selectedShapes[1];

            Drawing.PointF[] allPointsOfFirstRef = Graphics.GetRealCoordinates(firstRef);
            Drawing.PointF[] allPointsOfLastRef = Graphics.GetRealCoordinates(lastRef);
            Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);

            var horizontalDistance = Graphics.LeftMostPoint(allPointsOfLastRef).X - Graphics.RightMostPoint(allPointsOfFirstRef).X;
            var verticalDistance = Graphics.TopMostPoint(allPointsOfLastRef).Y - Graphics.BottomMostPoint(allPointsOfFirstRef).Y;

            var spaceBetweenShapes = horizontalDistance;

            for (int i = 2; i < shapeCount; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                var shapeWidth = Graphics.RealWidth(allPoints);
                spaceBetweenShapes -= shapeWidth;
            }

            // TODO: guard against spaceBetweenShapes < 0

            spaceBetweenShapes /= (shapeCount-1);
            
            for (int i = 2; i < shapeCount; i++)
            {
                Shape currShape = selectedShapes[i];
                Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(currShape);
                Drawing.PointF leftMost = Graphics.LeftMostPoint(allPoints);
                refShape = selectedShapes[i - 1];
                allPointsOfRef = Graphics.GetRealCoordinates(refShape);

                Drawing.PointF rightMostRef = Graphics.RightMostPoint(allPointsOfRef);
                currShape.IncrementLeft(rightMostRef.X - leftMost.X + spaceBetweenShapes);
            }

            spaceBetweenShapes = verticalDistance;
            for (int i = 2; i < shapeCount; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                var shapeHeight = Graphics.RealHeight(allPoints);
                spaceBetweenShapes -= shapeHeight;
            }

            // TODO: guard against spaceBetweenShapes < 0

            spaceBetweenShapes /= shapeCount;

            for (int i = 2; i < shapeCount; i++)
            {
                Shape currShape = selectedShapes[i];
                Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(currShape);
                Drawing.PointF topMost = Graphics.TopMostPoint(allPoints);
                refShape = selectedShapes[i - 1];
                allPointsOfRef = Graphics.GetRealCoordinates(refShape);

                Drawing.PointF lowestRef = Graphics.BottomMostPoint(allPointsOfRef);
                currShape.IncrementTop(lowestRef.Y - topMost.Y + spaceBetweenShapes);
            }
        }
        #endregion

        #endregion

        #region Util


        #endregion
    }
}
