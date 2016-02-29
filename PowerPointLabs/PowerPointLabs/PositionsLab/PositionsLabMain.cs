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
using AutoShape = Microsoft.Office.Core.MsoAutoShapeType;
using System.Diagnostics;
using Drawing = System.Drawing;

namespace PowerPointLabs.PositionsLab
{
    class PositionsLabMain
    {
        private static bool isInit = false;
        private const float epsilon = 0.00001f;
        private const float ROTATE_LEFT = 90f;
        private const float ROTATE_RIGHT = 270f;
        private const float ROTATE_UP = 0f;
        private const float ROTATE_DOWN = 180f;
        private const int NONE = -1;
        private const int RIGHT = 0;
        private const int DOWN = 1;
        private const int LEFT = 2;
        private const int UP = 3;
        private const int LEFTORRIGHT = 4;
        private const int UPORDOWN = 5;

        //For Grid
        public const int ALIGN_LEFT = 0;
        public const int ALIGN_CENTER = 1;
        public const int ALIGN_RIGHT = 2;

        private static Dictionary<MsoAutoShapeType, float> shapeDefaultUpAngle;
        private static bool _alignUseSlideAsReference = false;
        private static bool _distributeUseSlideAsReference = false;

        #region API

        #region Class Methods

        /// <summary>
        /// Tells the Positions Lab to use the slide as the reference point for Align methods
        /// </summary>
        public static void AlignReferToSlide()
        {
            _alignUseSlideAsReference = true;
        }

        /// <summary>
        /// Tells the Positions Lab to use first selected shape as reference shape for Align methods
        /// </summary>
        public static void AlignReferToShape()
        {
            _alignUseSlideAsReference = false;
        }

        /// <summary>
        /// Tells the Position Lab to use the slide as the reference point for Distribute methods
        /// </summary>
        public static void DistributeReferToSlide()
        {
            _distributeUseSlideAsReference = true;
        }

        /// <summary>
        /// Tells the Positions Lab to use first selected shape as reference shape for Distribute methods
        /// </summary>
        public static void DistributeReferToShape()
        {
            _distributeUseSlideAsReference = false;
        }

        #endregion

        #region Align
        public static void AlignLeft(List<Shape> selectedShapes)
        {
            if (_alignUseSlideAsReference)
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

                Shape refShape = selectedShapes[0];
                Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);
                Drawing.PointF leftMostRef = Graphics.LeftMostPoint(allPointsOfRef);

                for (int i = 1; i < selectedShapes.Count; i++)
                {
                    Shape s = selectedShapes[i];
                    Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                    Drawing.PointF leftMost = Graphics.LeftMostPoint(allPoints);
                    s.IncrementLeft(leftMostRef.X - leftMost.X);
                }
            }
        }

        public static void AlignRight(List<Shape> selectedShapes, float slideWidth)
        {
            if (_alignUseSlideAsReference)
            {
                foreach (Shape s in selectedShapes)
                {
                    Drawing.PointF[] allPointsOfShape = Graphics.GetRealCoordinates(s);
                    Drawing.PointF leftMost = Graphics.LeftMostPoint(allPointsOfShape);
                    var shapeWidth = Graphics.RealWidth(allPointsOfShape);
                    s.IncrementLeft(slideWidth - leftMost.X - shapeWidth);
                }
            }
            else
            {
                if (selectedShapes.Count < 2)
                {
                    //Error
                    return;
                }

                Shape refShape = selectedShapes[0];
                Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);
                Drawing.PointF rightMostRef = Graphics.RightMostPoint(allPointsOfRef);

                for (int i = 1; i < selectedShapes.Count; i++)
                {
                    Shape s = selectedShapes[i];
                    Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                    Drawing.PointF rightMost = Graphics.RightMostPoint(allPoints);
                    s.IncrementLeft(rightMostRef.X - rightMost.X);
                }
            }
        }

        public static void AlignTop(List<Shape> selectedShapes)
        {
            if (_alignUseSlideAsReference)
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

                Shape refShape = selectedShapes[0];
                Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);
                Drawing.PointF topMostRef = Graphics.TopMostPoint(allPointsOfRef);

                for (int i = 1; i < selectedShapes.Count; i++)
                {
                    Shape s = selectedShapes[i];
                    Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                    Drawing.PointF topMost = Graphics.TopMostPoint(allPoints);
                    s.IncrementTop(topMostRef.Y - topMost.Y);
                }
            }
        }

        public static void AlignBottom(List<Shape> selectedShapes, float slideHeight)
        {
            if (_alignUseSlideAsReference)
            {
                foreach (Shape s in selectedShapes)
                {
                    Drawing.PointF[] allPointsOfShape = Graphics.GetRealCoordinates(s);
                    Drawing.PointF topMost = Graphics.TopMostPoint(allPointsOfShape);
                    var shapeHeight = Graphics.RealHeight(allPointsOfShape);
                    s.IncrementTop(slideHeight - topMost.Y - shapeHeight);
                }
            }
            else
            {
                if (selectedShapes.Count < 2)
                {
                    //Error
                    return;
                }

                Shape refShape = selectedShapes[0];
                Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);
                Drawing.PointF lowestRef = Graphics.BottomMostPoint(allPointsOfRef);

                for (int i = 1; i < selectedShapes.Count; i++)
                {
                    Shape s = selectedShapes[i];
                    Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                    Drawing.PointF lowest = Graphics.BottomMostPoint(allPoints);
                    s.IncrementTop(lowestRef.Y - lowest.Y);
                }
            }
        }

        public static void AlignMiddle(List<Shape> selectedShapes, float slideHeight)
        {
            if (_alignUseSlideAsReference)
            {
                foreach (Shape s in selectedShapes)
                {
                    Drawing.PointF[] allPointsOfShape = Graphics.GetRealCoordinates(s);
                    Drawing.PointF topMost = Graphics.TopMostPoint(allPointsOfShape);
                    var shapeHeight = Graphics.RealHeight(allPointsOfShape);
                    s.IncrementTop(slideHeight/2 - topMost.Y - shapeHeight/2);
                }
            }
            else
            {
                if (selectedShapes.Count < 2)
                {
                    //Error
                    return;
                }

                Shape refShape = selectedShapes[0];
                Drawing.PointF originRef = Graphics.GetCenterPoint(refShape);

                for (int i = 1; i < selectedShapes.Count; i++)
                {
                    Shape s = selectedShapes[i];
                    Drawing.PointF origin = Graphics.GetCenterPoint(s);
                    s.IncrementTop(originRef.Y - origin.Y);
                }
            }
        }

        public static void AlignCenter(List<Shape> selectedShapes, float slideWidth, float slideHeight)
        {
            if (_alignUseSlideAsReference)
            {
                foreach (Shape s in selectedShapes)
                {
                    Drawing.PointF[] allPointsOfShape = Graphics.GetRealCoordinates(s);
                    Drawing.PointF topMost = Graphics.TopMostPoint(allPointsOfShape);
                    Drawing.PointF leftMost = Graphics.LeftMostPoint(allPointsOfShape);
                    var shapeHeight = Graphics.RealHeight(allPointsOfShape);
                    var shapeWidth = Graphics.RealWidth(allPointsOfShape);
                    s.IncrementTop(slideHeight/2 - topMost.Y - shapeHeight/2);
                    s.IncrementLeft(slideWidth/2 - leftMost.X - shapeWidth/2);
                }
            }
            else
            {
                if (selectedShapes.Count < 2)
                {
                    //Error
                    return;
                }

                Shape refShape = selectedShapes[0];
                Drawing.PointF originRef = Graphics.GetCenterPoint(refShape);

                for (int i = 1; i < selectedShapes.Count; i++)
                {
                    Shape s = selectedShapes[i];
                    Drawing.PointF origin = Graphics.GetCenterPoint(s);
                    s.IncrementLeft(originRef.X - origin.X);
                    s.IncrementTop(originRef.Y - origin.Y);
                }
            }
        }

        #endregion

        #region Snap
        public static void SnapVertical(List<Shape> selectedShapes)
        {
            if (!isInit)
            {
                Init();
                isInit = true;
            }

            for (int i = 0; i < selectedShapes.Count; i++)
            {
                SnapShapeVertical(selectedShapes[i]);
            }
        }

        public static void SnapHorizontal(List<Shape> selectedShapes)
        {
            if (!isInit)
            {
                Init();
                isInit = true;
            }

            for (int i = 0; i < selectedShapes.Count; i++)
            {
                SnapShapeHorizontal(selectedShapes[i]);
            }
        }

        public static void SnapAway(List<Shape> shapes)
        {
            if (!isInit)
            {
                Init();
                isInit = true;
            }

            if (shapes.Count <= 1)
            {
                return;
            }

            Drawing.PointF refShapeCenter = Graphics.GetCenterPoint(shapes[0]);
            bool isAllSameDir = true;
            int lastDir = -1;

            for (int i = 1; i < shapes.Count; i++)
            {
                Shape shape = shapes[i];
                Drawing.PointF shapeCenter = Graphics.GetCenterPoint(shape);
                float angle = (float)AngleBetweenTwoPoints(refShapeCenter, shapeCenter);

                int dir = GetDirectionWRTRefShape(shape, angle);

                if (i == 1)
                {
                    lastDir = dir;
                }

                if (!IsSameDirection(lastDir, dir))
                {
                    isAllSameDir = false;
                    break;
                }

                //only maintain in one direction instead of dual direction
                if (dir < LEFTORRIGHT)
                {
                    lastDir = dir; 
                }
            }

            if (!isAllSameDir || lastDir == NONE)
            {
                lastDir = 0;
            }
            else
            {
                lastDir++;
            }

            for (int i = 1; i < shapes.Count; i++)
            {
                Shape shape = shapes[i];
                Drawing.PointF shapeCenter = Graphics.GetCenterPoint(shape);
                float angle = (float) AngleBetweenTwoPoints(refShapeCenter, shapeCenter);

                float defaultUpAngle = 0;
                bool hasDefaultDirection = shapeDefaultUpAngle.TryGetValue(shape.AutoShapeType, out defaultUpAngle);

                if (hasDefaultDirection)
                {
                    shape.Rotation = (defaultUpAngle + angle) + lastDir * 90;
                }
                else
                {
                    if (IsVertical(shape))
                    {
                        shape.Rotation = angle + lastDir * 90;
                    }
                    else
                    {
                        shape.Rotation = (angle - 90) + lastDir * 90;
                    }
                }
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
        public static void AdjoinHorizontal(List<Shape> selectedShapes)
        {
            if (selectedShapes.Count < 2)
            {
                //Error
                return;
            }

            Shape refShape = selectedShapes[0];
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

        public static void AdjoinVertical(List<Shape> selectedShapes)
        {
            if (selectedShapes.Count < 2)
            {
                //Error
                return;
            }

            Shape refShape = selectedShapes[0];
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
        public static void Swap(List<Shape> selectedShapes)
        {
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
        public static void DistributeHorizontal(List<Shape> selectedShapes, float slideWidth)
        {

            if (_distributeUseSlideAsReference)
            {
                // TODO: Implement distribute horizontal with slide as reference
            }

            var shapeCount = selectedShapes.Count;

            Shape refShape = selectedShapes[0];
            Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);
            Drawing.PointF rightMostRef;

            if (_alignUseSlideAsReference)
            {
                var horizontalDistanceInRef = slideWidth;
                var spaceBetweenShapes = horizontalDistanceInRef;

                foreach (Shape s in selectedShapes)
                {
                    Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                    var shapeWidth = Graphics.RealWidth(allPoints);
                    spaceBetweenShapes -= shapeWidth;
                }

                // TODO: guard against spaceBetweenShapes < 0

                spaceBetweenShapes /= shapeCount + 1;

                for (int i = 0; i < shapeCount; i++)
                {
                    Shape currShape = selectedShapes[i];
                    Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(currShape);
                    Drawing.PointF leftMost = Graphics.LeftMostPoint(allPoints);
                    if (i == 0)
                    {
                        currShape.IncrementLeft(spaceBetweenShapes - leftMost.X);
                    }
                    else
                    {
                        refShape = selectedShapes[i - 1];
                        allPointsOfRef = Graphics.GetRealCoordinates(refShape);
                        rightMostRef = Graphics.RightMostPoint(allPointsOfRef);
                        currShape.IncrementLeft(rightMostRef.X - leftMost.X + spaceBetweenShapes);
                    }
                }
            }
            else
            {
                if (shapeCount < 2)
                {
                    //Error
                    return;
                }

                var horizontalDistanceInRef = Graphics.RealWidth(allPointsOfRef);
                var spaceBetweenShapes = horizontalDistanceInRef;

                for (int i = 1; i < shapeCount; i++)
                {
                    Shape s = selectedShapes[i];
                    Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                    var shapeWidth = Graphics.RealWidth(allPoints);
                    spaceBetweenShapes -= shapeWidth;
                }

                // TODO: guard against spaceBetweenShapes < 0

                spaceBetweenShapes /= shapeCount;

                for (int i = 1; i < shapeCount; i++)
                {
                    Shape currShape = selectedShapes[i];
                    Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(currShape);
                    Drawing.PointF leftMost = Graphics.LeftMostPoint(allPoints);
                    refShape = selectedShapes[i - 1];
                    allPointsOfRef = Graphics.GetRealCoordinates(refShape);

                    if (i == 1)
                    {
                        Drawing.PointF leftMostRef = Graphics.LeftMostPoint(allPointsOfRef);
                        currShape.IncrementLeft(leftMostRef.X - leftMost.X + spaceBetweenShapes);
                    }
                    else
                    {
                        rightMostRef = Graphics.RightMostPoint(allPointsOfRef);
                        currShape.IncrementLeft(rightMostRef.X - leftMost.X + spaceBetweenShapes);
                    }
                }
            }
        } 

        public static void DistributeVertical(List<Shape> selectedShapes, float slideHeight)
        {
            var shapeCount = selectedShapes.Count;

            Shape refShape = selectedShapes[0];
            Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);
            Drawing.PointF lowestRef;

            if (_alignUseSlideAsReference)
            {
                var verticalDistanceInRef = slideHeight;
                var spaceBetweenShapes = verticalDistanceInRef;

                foreach (Shape s in selectedShapes)
                {
                    Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                    var shapeHeight = Graphics.RealHeight(allPoints);
                    spaceBetweenShapes -= shapeHeight;
                }

                // TODO: guard against spaceBetweenShapes < 0

                spaceBetweenShapes /= shapeCount + 1;

                for (int i = 0; i < shapeCount; i++)
                {
                    Shape currShape = selectedShapes[i];
                    Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(currShape);
                    Drawing.PointF topMost = Graphics.TopMostPoint(allPoints);
                    if (i == 0)
                    {
                        currShape.IncrementTop(spaceBetweenShapes - topMost.Y);
                    }
                    else
                    {
                        refShape = selectedShapes[i - 1];
                        allPointsOfRef = Graphics.GetRealCoordinates(refShape);
                        lowestRef = Graphics.BottomMostPoint(allPointsOfRef);
                        currShape.IncrementTop(lowestRef.Y - topMost.Y + spaceBetweenShapes);
                    }
                }
            }
            else
            {
                if (shapeCount < 2)
                {
                    //Error
                    return;
                }

                var verticalDistanceInRef = Graphics.RealHeight(allPointsOfRef);
                var spaceBetweenShapes = verticalDistanceInRef;

                for (int i = 1; i < shapeCount; i++)
                {
                    Shape s = selectedShapes[i];
                    Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                    var shapeHeight = Graphics.RealHeight(allPoints);
                    spaceBetweenShapes -= shapeHeight;
                }

                // TODO: guard against spaceBetweenShapes < 0

                spaceBetweenShapes /= shapeCount;

                for (int i = 1; i < shapeCount; i++)
                {
                    Shape currShape = selectedShapes[i];
                    Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(currShape);
                    Drawing.PointF topMost = Graphics.TopMostPoint(allPoints);
                    refShape = selectedShapes[i - 1];
                    allPointsOfRef = Graphics.GetRealCoordinates(refShape);

                    if (i == 1)
                    {
                        Drawing.PointF topMostRef = Graphics.TopMostPoint(allPointsOfRef);
                        currShape.IncrementTop(topMostRef.Y - topMost.Y + spaceBetweenShapes);
                    }
                    else
                    {
                        lowestRef = Graphics.BottomMostPoint(allPointsOfRef);
                        currShape.IncrementTop(lowestRef.Y - topMost.Y + spaceBetweenShapes);
                    }
                }
            }
        }

        public static void DistributeCenter(List<Shape> selectedShapes, float slideWidth, float slideHeight)
        {
            DistributeHorizontal(selectedShapes, slideWidth);
            DistributeVertical(selectedShapes, slideHeight);
        }

        public static void DistributeShapes(List<Shape> selectedShapes)
        {
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

            Shape firstRef = selectedShapes[0];
            Shape lastRef = selectedShapes[selectedShapes.Count - 1];
            Shape refShape = selectedShapes[0];

            Drawing.PointF[] allPointsOfFirstRef = Graphics.GetRealCoordinates(firstRef);
            Drawing.PointF[] allPointsOfLastRef = Graphics.GetRealCoordinates(lastRef);
            Drawing.PointF[] allPointsOfRef = Graphics.GetRealCoordinates(refShape);

            var horizontalDistance = Graphics.LeftMostPoint(allPointsOfLastRef).X - Graphics.RightMostPoint(allPointsOfFirstRef).X;
            var verticalDistance = Graphics.TopMostPoint(allPointsOfLastRef).Y - Graphics.BottomMostPoint(allPointsOfFirstRef).Y;

            var spaceBetweenShapes = horizontalDistance;

            for (int i = 1; i < shapeCount - 1; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                var shapeWidth = Graphics.RealWidth(allPoints);
                spaceBetweenShapes -= shapeWidth;
            }

            // TODO: guard against spaceBetweenShapes < 0

            spaceBetweenShapes /= (shapeCount-1);
            
            for (int i = 1; i < shapeCount - 1; i++)
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
            for (int i = 1; i < shapeCount - 1; i++)
            {
                Shape s = selectedShapes[i];
                Drawing.PointF[] allPoints = Graphics.GetRealCoordinates(s);
                var shapeHeight = Graphics.RealHeight(allPoints);
                spaceBetweenShapes -= shapeHeight;
            }

            // TODO: guard against spaceBetweenShapes < 0

            spaceBetweenShapes /= shapeCount;

            for (int i = 1; i < shapeCount - 1; i++)
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

        public static void DistributeGrid(List<Shape> selectedShapes, int rowLength, int colLength, float marginTop, float marginBottom, float marginLeft, float marginRight, int alignment)
        {
            Drawing.PointF refPoint = Graphics.GetCenterPoint(selectedShapes[0]);

            List<PPShape> allShapes = new List<PPShape>();
            int numShapes = selectedShapes.Count;

            for (int i = 0; i < numShapes; i++)
            {
                allShapes.Add(new PPShape(selectedShapes[i]));
            }

            float[] rowDifferences = GetLongestWidthsOfRows(allShapes, rowLength);
            float[] colDifferences = GetLongestHeightsOfCols(allShapes, rowLength, colLength);

            float posX = refPoint.X;
            float posY = refPoint.Y;
            int numIndicesToSkip = IndicesToSkip(numShapes, rowLength, alignment);
            int differenceIndex = 0;

            for (int i = 0; i < colLength; i++)
            {
                for (int j = 0; j < rowLength; j++)
                {
                    int index = i * rowLength + j;

                    if (index >= numShapes)
                    {
                        break;
                    }

                    Shape currentShape = selectedShapes[index];
                    Drawing.PointF center = Graphics.GetCenterPoint(currentShape);
                    currentShape.IncrementLeft(posX - center.X);
                    currentShape.IncrementTop(posY - center.Y);
                    
                    if (differenceIndex < rowLength - 1)
                    {
                        posX += (rowDifferences[differenceIndex] / 2 + rowDifferences[differenceIndex + 1] / 2 + marginLeft + marginRight);
                    }

                    differenceIndex++;
                }

                posX = refPoint.X;

                if (i < colLength - 1)
                {
                    differenceIndex = 0;
                    posY += (colDifferences[i] / 2 + colDifferences[i + 1] / 2 + marginTop + marginBottom);
                }

                if (i == colLength - 2)
                {
                    differenceIndex = numIndicesToSkip;
                    for (int k = 0; k < numIndicesToSkip; k++)
                    {
                        posX += (rowDifferences[k] / 2 + rowDifferences[k + 1] / 2 + marginLeft + marginRight);
                    }
                }
            }
        }
        #endregion

        #endregion

        #region Util
        public static double AngleBetweenTwoPoints(System.Drawing.PointF refPoint, System.Drawing.PointF pt)
        {
            double angle = Math.Atan((pt.Y - refPoint.Y) / (pt.X - refPoint.X)) * 180 / Math.PI;

            if (pt.X - refPoint.X > 0)
            {
                angle = 90 + angle;
            }
            else
            {
                angle = 270 + angle;
            }

            return angle;
        }

        public static bool NearlyEqual(float a, float b, float epsilon)
        {
            float absA = Math.Abs(a);
            float absB = Math.Abs(b);
            float diff = Math.Abs(a - b);

            if (a == b)
            { // shortcut, handles infinities
                return true;
            }
            else if (a == 0 || b == 0 || diff < float.Epsilon)
            {
                // a or b is zero or both are extremely close to it
                // relative error is less meaningful here
                return diff < epsilon;
            }
            else
            { // use relative error
                return diff / (absA + absB) < epsilon;
            }
        }

        private static int GetDirectionWRTRefShape(Shape shape, float angleFromRefShape)
        {
            float defaultUpAngle = -1;
            bool hasDefaultDirection = shapeDefaultUpAngle.TryGetValue(shape.AutoShapeType, out defaultUpAngle);

            if (shape.AutoShapeType == AutoShape.msoShapeLightningBolt)
            {
                Debug.WriteLine("defaultDir: " + hasDefaultDirection);
                Debug.WriteLine("defaultAngle: " + defaultUpAngle);
            }

            if (!hasDefaultDirection)
            {
                if (IsVertical(shape))
                {
                    defaultUpAngle = 0;
                }
                else
                {
                    defaultUpAngle = 90;
                }
            }

            float angle = AddAngles(angleFromRefShape, defaultUpAngle);
            float diff = SubtractAngles(shape.Rotation, angle);
            float phaseInFloat = diff / 90;

            if (shape.AutoShapeType == AutoShape.msoShapeLightningBolt)
            {
                Debug.WriteLine("angle: " + angle);
                Debug.WriteLine("diff: " + diff);
                Debug.WriteLine("phaseInFloat: " + defaultUpAngle);
                Debug.WriteLine("equal: " + NearlyEqual(phaseInFloat, (float)Math.Round(phaseInFloat), epsilon));
            }

            if (!NearlyEqual(phaseInFloat, (float)Math.Round(phaseInFloat), epsilon))
            {
                return NONE;
            }

            int phase = (int)Math.Round(phaseInFloat);

            if (!hasDefaultDirection)
            {
                if (phase == LEFT || phase == RIGHT)
                {
                    return LEFTORRIGHT;
                }

                return UPORDOWN;
            }

            return phase;
        }

        private static bool IsSameDirection(int a, int b)
        {
            if (a == b) return true;
            if (a == LEFTORRIGHT) return b == LEFT || b == RIGHT;
            if (b == LEFTORRIGHT) return a == LEFT || a == RIGHT;
            if (a == UPORDOWN) return b == UP || b == DOWN;
            if (b == UPORDOWN) return a == UP || a == DOWN;

            return false;
       }

        public static float AddAngles(float a, float b)
        {
            return (a + b) % 360;
        }

        public static float SubtractAngles(float a, float b)
        {
            float diff = a - b;
            if (diff < 0)
            {
                return 360 + diff;
            }

            return diff;
        }

        public static float[] GetLongestWidthsOfRows(List<PPShape> shapes, int rowLength)
        {
            float[] longestWidths = new float[rowLength];

            for (int i = 0; i < shapes.Count; i++)
            {
                int longestRowIndex = i % rowLength;
                if (longestWidths[longestRowIndex] < shapes[i].AbsoluteWidth)
                {
                    longestWidths[longestRowIndex] = shapes[i].AbsoluteWidth;
                }
            }

            return longestWidths;
        }

        public static float[] GetLongestHeightsOfCols(List<PPShape> shapes, int rowLength, int colLength)
        {
            float[] longestHeights = new float[colLength];

            for (int i = 0; i < shapes.Count; i++)
            {
                int longestHeightIndex = i / rowLength;
                if (longestHeights[longestHeightIndex] < shapes[i].AbsoluteHeight)
                {
                    longestHeights[longestHeightIndex] = shapes[i].AbsoluteHeight;
                }
            }

            return longestHeights;
        }

        public static int IndicesToSkip(int totalSelectedShapes, int rowLength, int alignment)
        {
            if (alignment == ALIGN_LEFT)
            {
                return 0;
            }

            int numOfShapesInLastRow = totalSelectedShapes % rowLength;

            if (alignment == ALIGN_RIGHT)
            {
                return rowLength - numOfShapesInLastRow;
            }

            if (alignment == ALIGN_CENTER)
            {
                int difference = rowLength - numOfShapesInLastRow;
                return difference / 2;
            }

            return 0;
        }

        private static void Init()
        {
            shapeDefaultUpAngle = new Dictionary<MsoAutoShapeType, float>();

            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftArrow, ROTATE_LEFT);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftRightArrow, ROTATE_LEFT);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftArrowCallout, ROTATE_LEFT);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftRightArrowCallout, ROTATE_LEFT);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCurvedLeftArrow, ROTATE_LEFT);

            shapeDefaultUpAngle.Add(AutoShape.msoShapeRightArrow, ROTATE_RIGHT);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeBentArrow, ROTATE_RIGHT);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeStripedRightArrow, ROTATE_RIGHT);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeNotchedRightArrow, ROTATE_RIGHT);
            shapeDefaultUpAngle.Add(AutoShape.msoShapePentagon, ROTATE_RIGHT);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeChevron, ROTATE_RIGHT);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeRightArrowCallout, ROTATE_RIGHT);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCurvedRightArrow, ROTATE_RIGHT);

            shapeDefaultUpAngle.Add(AutoShape.msoShapeUpArrow, ROTATE_UP);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeBentUpArrow, ROTATE_UP);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeUpDownArrow, ROTATE_UP);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftRightUpArrow, ROTATE_UP);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftUpArrow, ROTATE_UP);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeUpArrowCallout, ROTATE_UP);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCurvedUpArrow, ROTATE_UP);

            shapeDefaultUpAngle.Add(AutoShape.msoShapeDownArrow, ROTATE_DOWN);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeUTurnArrow, ROTATE_DOWN);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeDownArrowCallout, ROTATE_DOWN);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCurvedDownArrow, ROTATE_DOWN);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCircularArrow, ROTATE_DOWN);
        }

        #endregion
    }
}
