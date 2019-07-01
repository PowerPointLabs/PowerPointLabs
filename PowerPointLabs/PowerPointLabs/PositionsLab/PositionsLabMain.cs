using System;
using System.Collections.Generic;
using System.Linq;

using Microsoft.Office.Core;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

using AutoShape = Microsoft.Office.Core.MsoAutoShapeType;
using Drawing = System.Drawing;
using Office = Microsoft.Office.Core;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;

namespace PowerPointLabs.PositionsLab
{
    public class PositionsLabMain
    {
        private const float Epsilon = 0.00001f;
        private const float RotateLeft = 90f;
        private const float RotateRight = 270f;
        private const float RotateUp = 0f;
        private const float RotateDown = 180f;
        private const float threshold = 0.01f;
        private const int None = -1;
        private const int Right = 0;
        private const int Down = 1;
        private const int Left = 2;
        private const int Up = 3;
        private const int Leftorright = 4;
        private const int Upordown = 5;

        public class GridSpace
        {
            public float RowDifference { get; }
            public float ColDifference { get; }
            public GridSpace(float rowDifference, float colDifference)
            {
                RowDifference = rowDifference;
                ColDifference = colDifference;
            }
        }

        private class ShapeAngleInfo
        {
            public Shape Shape { get; private set; }
            public float Angle { get; set; }
            public float ShapeAngle { get; set; }

            public ShapeAngleInfo(Shape shape, float angle)
            {
                Shape = shape;
                Angle = angle;
            }

            public ShapeAngleInfo(Shape shape, float angle, float shapeAngle) : this(shape, angle)
            {
                ShapeAngle = shapeAngle;
            }
        }

        private static Dictionary<string, Drawing.PointF> prevSelectedShapes = new Dictionary<string, Drawing.PointF>();
        private static List<string> prevSortedShapeNames;

        // Adjoin Variables
        public static bool AlignShapesToBeAdjoined { get; private set; }

        private static Dictionary<MsoAutoShapeType, float> shapeDefaultUpAngle;

        #region API

        #region Class Methods

        /// <summary>
        /// Tells the Positions Lab to align the shapes that are to be adjoined
        /// </summary>
        public static void AdjoinWithAligning()
        {
            AlignShapesToBeAdjoined = true;
        }

        /// <summary>
        /// Tells the Positions Lab to not align the shapes that are to be adjoined
        /// </summary>
        public static void AdjoinWithoutAligning()
        {
            AlignShapesToBeAdjoined = false;
        }

        #endregion

        #region Align
        public static void AlignLeft(ShapeRange toAlign)
        {
            List<PPShape> selectedShapes = new List<PPShape>();

            switch (PositionsLabSettings.AlignReference)
            {
                case PositionsLabSettings.AlignReferenceObject.Slide:
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    foreach (PPShape s in selectedShapes)
                    {
                        s.IncrementLeft(-s.VisualLeft);
                    }
                    break;
                case PositionsLabSettings.AlignReferenceObject.SelectedShape:
                    if (toAlign.Count < 2)
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignLefts, MsoTriState.msoTrue);
                        break;
                    }

                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    PPShape refShape = selectedShapes[0];

                    for (int i = 1; i < selectedShapes.Count; i++)
                    {
                        PPShape s = selectedShapes[i];
                        s.IncrementLeft(refShape.VisualLeft - s.VisualLeft);
                    }
                    break;
                case PositionsLabSettings.AlignReferenceObject.PowerpointDefaults:

                    if (toAlign.Count == 1)
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignLefts, MsoTriState.msoTrue);
                    }
                    else
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignLefts, MsoTriState.msoFalse);
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static void AlignRight(ShapeRange toAlign, float slideWidth)
        {

            List<PPShape> selectedShapes = new List<PPShape>();

            switch (PositionsLabSettings.AlignReference)
            {
                case PositionsLabSettings.AlignReferenceObject.Slide:
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    foreach (PPShape s in selectedShapes)
                    {
                        s.IncrementLeft(slideWidth - s.VisualLeft - s.AbsoluteWidth);
                    }
                    break;
                case PositionsLabSettings.AlignReferenceObject.SelectedShape:
                    if (toAlign.Count < 2)
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignRights, MsoTriState.msoTrue);
                        break;
                    }

                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    PPShape refShape = selectedShapes[0];
                    float rightMostRefPoint = refShape.VisualLeft + refShape.AbsoluteWidth;

                    for (int i = 1; i < selectedShapes.Count; i++)
                    {
                        PPShape s = selectedShapes[i];
                        float rightMostPoint = s.VisualLeft + s.AbsoluteWidth;
                        s.IncrementLeft(rightMostRefPoint - rightMostPoint);
                    }
                    break;
                case PositionsLabSettings.AlignReferenceObject.PowerpointDefaults:
                    if (toAlign.Count == 1)
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignRights, MsoTriState.msoTrue);
                    }
                    else
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignRights, MsoTriState.msoFalse);
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static void AlignTop(ShapeRange toAlign)   
        {
            List<PPShape> selectedShapes = new List<PPShape>();

            switch (PositionsLabSettings.AlignReference)
            {
                case PositionsLabSettings.AlignReferenceObject.Slide:
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    foreach (PPShape s in selectedShapes)
                    {
                        s.IncrementTop(-s.VisualTop);
                    }
                    break;
                case PositionsLabSettings.AlignReferenceObject.SelectedShape:
                    if (toAlign.Count < 2)
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignTops, MsoTriState.msoTrue);
                        break;
                    }

                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    PPShape refShape = selectedShapes[0];

                    for (int i = 1; i < selectedShapes.Count; i++)
                    {
                        PPShape s = selectedShapes[i];
                        s.IncrementTop(refShape.VisualTop - s.VisualTop);
                    }
                    break;
                case PositionsLabSettings.AlignReferenceObject.PowerpointDefaults:
                    if (toAlign.Count == 1)
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignTops, MsoTriState.msoTrue);
                    }
                    else
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignTops, MsoTriState.msoFalse);
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static void AlignBottom(ShapeRange toAlign, float slideHeight)
        {
            List<PPShape> selectedShapes = new List<PPShape>();

            switch (PositionsLabSettings.AlignReference)
            {
                case PositionsLabSettings.AlignReferenceObject.Slide:
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    foreach (PPShape s in selectedShapes)
                    {
                        s.IncrementTop(slideHeight - s.VisualTop - s.AbsoluteHeight);
                    }
                    break;
                case PositionsLabSettings.AlignReferenceObject.SelectedShape:
                    if (toAlign.Count < 2)
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignBottoms, MsoTriState.msoTrue);
                        break;
                    }

                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    PPShape refShape = selectedShapes[0];
                    float lowestRefPoint = refShape.VisualTop + refShape.AbsoluteHeight;

                    for (int i = 1; i < selectedShapes.Count; i++)
                    {
                        PPShape s = selectedShapes[i];
                        float lowestPoint = s.VisualTop + s.AbsoluteHeight;
                        s.IncrementTop(lowestRefPoint - lowestPoint);
                    }
                    break;
                case PositionsLabSettings.AlignReferenceObject.PowerpointDefaults:
                    if (toAlign.Count == 1)
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignBottoms, MsoTriState.msoTrue);
                    }
                    else
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignBottoms, MsoTriState.msoFalse);
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static void AlignHorizontalCenter(ShapeRange toAlign, float slideHeight)
        {
            List<PPShape> selectedShapes = new List<PPShape>();

            switch (PositionsLabSettings.AlignReference)
            {
                case PositionsLabSettings.AlignReferenceObject.Slide:
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    foreach (PPShape s in selectedShapes)
                    {
                        s.IncrementTop(slideHeight / 2 - s.VisualTop - s.AbsoluteHeight / 2);
                    }
                    break;
                case PositionsLabSettings.AlignReferenceObject.SelectedShape:
                    if (toAlign.Count < 2)
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignMiddles, MsoTriState.msoTrue);
                        break;
                    }

                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    PPShape refShape = selectedShapes[0];

                    for (int i = 1; i < selectedShapes.Count; i++)
                    {
                        PPShape s = selectedShapes[i];
                        s.IncrementTop(refShape.VisualCenter.Y - s.VisualCenter.Y);
                    }
                    break;
                case PositionsLabSettings.AlignReferenceObject.PowerpointDefaults:
                    if (toAlign.Count == 1)
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignMiddles, MsoTriState.msoTrue);
                    }
                    else
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignMiddles, MsoTriState.msoFalse);
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static void AlignVerticalCenter(ShapeRange toAlign, float slideWidth)
        {
            List<PPShape> selectedShapes = new List<PPShape>();

            switch (PositionsLabSettings.AlignReference)
            {
                case PositionsLabSettings.AlignReferenceObject.Slide:
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    foreach (PPShape s in selectedShapes)
                    {
                        s.IncrementLeft(slideWidth / 2 - s.VisualLeft - s.AbsoluteWidth / 2);
                    }
                    break;
                case PositionsLabSettings.AlignReferenceObject.SelectedShape:
                    if (toAlign.Count < 2)
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignCenters, MsoTriState.msoTrue);
                        break;
                    }

                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    PPShape refShape = selectedShapes[0];

                    for (int i = 1; i < selectedShapes.Count; i++)
                    {
                        PPShape s = selectedShapes[i];
                        s.IncrementLeft(refShape.VisualCenter.X - s.VisualCenter.X);
                    }
                    break;
                case PositionsLabSettings.AlignReferenceObject.PowerpointDefaults:
                    if (toAlign.Count == 1)
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignCenters, MsoTriState.msoTrue);
                    }
                    else
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignCenters, MsoTriState.msoFalse);
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static void AlignCenter(ShapeRange toAlign, float slideHeight, float slideWidth)
        {
            List<PPShape> selectedShapes = new List<PPShape>();

            switch (PositionsLabSettings.AlignReference)
            {
                case PositionsLabSettings.AlignReferenceObject.Slide:
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    foreach (PPShape s in selectedShapes)
                    {
                        s.IncrementTop(slideHeight / 2 - s.VisualTop - s.AbsoluteHeight / 2);
                        s.IncrementLeft(slideWidth / 2 - s.VisualLeft - s.AbsoluteWidth / 2);
                    }
                    break;
                case PositionsLabSettings.AlignReferenceObject.SelectedShape:
                    if (toAlign.Count < 2)
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignMiddles, MsoTriState.msoTrue);
                        toAlign.Align(MsoAlignCmd.msoAlignCenters, MsoTriState.msoTrue);
                        break;
                    }
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    PPShape refShape = selectedShapes[0];

                    for (int i = 1; i < selectedShapes.Count; i++)
                    {
                        PPShape s = selectedShapes[i];
                        s.IncrementTop(refShape.VisualCenter.Y - s.VisualCenter.Y);
                        s.IncrementLeft(refShape.VisualCenter.X - s.VisualCenter.X);
                    }
                    break;
                case PositionsLabSettings.AlignReferenceObject.PowerpointDefaults:
                    if (toAlign.Count == 1)
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignMiddles, MsoTriState.msoTrue);
                        toAlign.Align(MsoAlignCmd.msoAlignCenters, MsoTriState.msoTrue);
                    }
                    else
                    {
                        toAlign.Align(MsoAlignCmd.msoAlignMiddles, MsoTriState.msoFalse);
                        toAlign.Align(MsoAlignCmd.msoAlignCenters, MsoTriState.msoFalse);
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static void AlignRadial(ShapeRange selectedShapes)
        {
            if (selectedShapes.Count < 3)
            {
                throw new Exception(PositionsLabText.ErrorFewerThanThreeSelection);
            }
                    
            Drawing.PointF origin = ShapeUtil.GetCenterPoint(selectedShapes[1]);
            Drawing.PointF refPoint = ShapeUtil.GetCenterPoint(selectedShapes[2]);
            double distance = DistanceBetweenTwoPoints(origin, refPoint);

            for (int i = 3; i <= selectedShapes.Count; i++)
            {
                Shape shape = selectedShapes[i];
                Drawing.PointF point = ShapeUtil.GetCenterPoint(shape);
                double currentDistance = DistanceBetweenTwoPoints(origin, point);
                double proportion = (currentDistance - distance) / currentDistance;

                shape.IncrementLeft((float)((origin.X - point.X) * proportion));
                shape.IncrementTop((float)((origin.Y - point.Y) * proportion));
            }
        }

        #endregion

        #region Adjoin
        public static void AdjoinHorizontal(List<PPShape> selectedShapes)
        {
            if (selectedShapes.Count < 2)
            {
                throw new Exception(PositionsLabText.ErrorFewerThanTwoSelection);
            }

            PPShape refShape = selectedShapes[0];
            List<PPShape> sortedShapes = ShapeUtil.SortShapesByLeft(selectedShapes);
            int refShapeIndex = sortedShapes.IndexOf(refShape);

            float mostLeft = refShape.VisualLeft;
            //For all shapes left of refShape, adjoin them from closest to refShape
            for (int i = refShapeIndex - 1; i >= 0; i--)
            {
                PPShape neighbour = sortedShapes[i];
                float rightOfNeighbour = neighbour.VisualLeft + neighbour.AbsoluteWidth;
                neighbour.IncrementLeft(mostLeft - rightOfNeighbour);
                if (AlignShapesToBeAdjoined)
                {
                    neighbour.IncrementTop(refShape.VisualCenter.Y - neighbour.VisualCenter.Y);
                }

                mostLeft = mostLeft - neighbour.AbsoluteWidth;
            }

            float mostRight = refShape.VisualLeft + refShape.AbsoluteWidth;
            //For all shapes right of refShape, adjoin them from closest to refShape
            for (int i = refShapeIndex + 1; i < sortedShapes.Count; i++)
            {
                PPShape neighbour = sortedShapes[i];
                neighbour.IncrementLeft(mostRight - neighbour.VisualLeft);
                if (AlignShapesToBeAdjoined)
                {
                    neighbour.IncrementTop(refShape.VisualCenter.Y - neighbour.VisualCenter.Y);
                }

                mostRight = mostRight + neighbour.AbsoluteWidth;
            }
        }

        public static void AdjoinVertical(List<PPShape> selectedShapes)
        {
            if (selectedShapes.Count < 2)
            {
                throw new Exception(PositionsLabText.ErrorFewerThanTwoSelection);
            }

            PPShape refShape = selectedShapes[0];
            List<PPShape> sortedShapes = ShapeUtil.SortShapesByTop(selectedShapes);
            int refShapeIndex = sortedShapes.IndexOf(refShape);

            float mostTop = refShape.VisualTop;
            //For all shapes above refShape, adjoin them from closest to refShape
            for (int i = refShapeIndex - 1; i >= 0; i--)
            {
                PPShape neighbour = sortedShapes[i];
                float bottomOfNeighbour = neighbour.VisualTop + neighbour.AbsoluteHeight;
                if (AlignShapesToBeAdjoined)
                {
                    neighbour.IncrementLeft(refShape.VisualCenter.X - neighbour.VisualCenter.X);
                }
                neighbour.IncrementTop(mostTop - bottomOfNeighbour);

                mostTop = mostTop - neighbour.AbsoluteHeight;
            }

            float lowest = refShape.VisualTop + refShape.AbsoluteHeight;
            //For all shapes right of refShape, adjoin them from closest to refShape
            for (int i = refShapeIndex + 1; i < sortedShapes.Count; i++)
            {
                PPShape neighbour = sortedShapes[i];
                if (AlignShapesToBeAdjoined)
                {
                    neighbour.IncrementLeft(refShape.VisualCenter.X - neighbour.VisualCenter.X);
                }
                neighbour.IncrementTop(lowest - neighbour.VisualTop);

                lowest = lowest + neighbour.AbsoluteHeight;
            }
        }
        #endregion

        #region Distribute

        public static void DistributeHorizontal(List<PPShape> selectedShapes, float slideWidth)
        {
            bool isSlide = PositionsLabSettings.DistributeReference == PositionsLabSettings.DistributeReferenceObject.Slide;
            bool isFirstShape = PositionsLabSettings.DistributeReference == PositionsLabSettings.DistributeReferenceObject.FirstShape;
            bool isExtremeShape = PositionsLabSettings.DistributeReference == PositionsLabSettings.DistributeReferenceObject.ExtremeShapes;
            bool isFirstTwoShapes = PositionsLabSettings.DistributeReference == PositionsLabSettings.DistributeReferenceObject.FirstTwoShapes;
            bool isObjectCenter = PositionsLabSettings.DistributeSpaceReference == PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            bool isObjectBoundary = PositionsLabSettings.DistributeSpaceReference == PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary;

            List<PPShape> shapesToDistribute;
            PPShape refShape;
            float referenceWidth, spaceBetweenShapes, startingPoint, rightMostRef, totalShapeWidth = 0;

            if (isSlide && isObjectBoundary)
            {
                if (selectedShapes.Count < 1)
                {
                    throw new Exception(PositionsLabText.ErrorNoSelection);
                }

                startingPoint = 0;
                referenceWidth = slideWidth;
                shapesToDistribute = ShapeUtil.SortShapesByLeft(selectedShapes);

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeWidth += s.AbsoluteWidth;
                }

                // Check if need leading and trailing space
                if (totalShapeWidth > referenceWidth)
                {
                    spaceBetweenShapes = (referenceWidth - totalShapeWidth) / (shapesToDistribute.Count - 1);
                }
                else
                {
                    spaceBetweenShapes = (referenceWidth - totalShapeWidth) / (shapesToDistribute.Count + 1);
                }

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];

                    // Check if need leading and trailing space
                    if (i == 0)
                    {
                        if (spaceBetweenShapes < 0)
                        {
                            currShape.IncrementLeft(startingPoint - currShape.VisualLeft);
                        }
                        else
                        {
                            currShape.IncrementLeft(startingPoint - currShape.VisualLeft + spaceBetweenShapes);
                        }
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        rightMostRef = refShape.VisualLeft + refShape.AbsoluteWidth;
                        currShape.IncrementLeft(rightMostRef - currShape.VisualLeft + spaceBetweenShapes);
                    }
                }
                return;
            }

            if (isSlide && isObjectCenter)
            {
                if (selectedShapes.Count < 1)
                {
                    throw new Exception(PositionsLabText.ErrorNoSelection);
                }

                startingPoint = 0;
                referenceWidth = slideWidth;
                shapesToDistribute = ShapeUtil.SortShapesByLeft(selectedShapes);

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeWidth += s.AbsoluteWidth;
                }

                if (totalShapeWidth > referenceWidth)
                {
                    PPShape leftMostShape = shapesToDistribute[0];
                    PPShape rightMostShape = shapesToDistribute[shapesToDistribute.Count - 1];

                    leftMostShape.IncrementLeft(startingPoint - leftMostShape.VisualLeft);
                    rightMostShape.IncrementLeft(referenceWidth - rightMostShape.VisualLeft - rightMostShape.AbsoluteWidth);

                    referenceWidth = referenceWidth - (leftMostShape.AbsoluteWidth / 2) - (rightMostShape.AbsoluteWidth / 2);

                    spaceBetweenShapes = referenceWidth/(shapesToDistribute.Count - 1);
                    startingPoint = leftMostShape.VisualCenter.X;
                    shapesToDistribute.RemoveAt(shapesToDistribute.Count - 1);
                    shapesToDistribute.RemoveAt(0);
                }
                else
                {
                    spaceBetweenShapes = referenceWidth/(shapesToDistribute.Count + 1);
                }

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];
                    if (i == 0)
                    {
                        currShape.IncrementLeft(startingPoint - currShape.VisualCenter.X + spaceBetweenShapes);
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        rightMostRef = refShape.ActualCenter.X;
                        currShape.IncrementLeft(rightMostRef - currShape.VisualCenter.X + spaceBetweenShapes);
                    }
                }
                return;
            }

            if (isFirstShape && isObjectBoundary)
            {
                if (selectedShapes.Count < 2)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanTwoSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualLeft;
                referenceWidth = shapesToDistribute[0].AbsoluteWidth;
                shapesToDistribute.RemoveAt(0);
                shapesToDistribute = ShapeUtil.SortShapesByLeft(shapesToDistribute);

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeWidth += s.AbsoluteWidth;
                }

                // Check if need leading and trailing space
                if (totalShapeWidth > referenceWidth)
                {
                    spaceBetweenShapes = (referenceWidth - totalShapeWidth) / (shapesToDistribute.Count - 1);
                }
                else
                {
                    spaceBetweenShapes = (referenceWidth - totalShapeWidth) / (shapesToDistribute.Count + 1);
                } 

                for (int i =0; i <shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];

                    // Check if need leading and trailing space
                    if (i==0)
                    {
                        if (spaceBetweenShapes < 0)
                        {
                            currShape.IncrementLeft(startingPoint - currShape.VisualLeft);
                        }
                        else
                        {
                            currShape.IncrementLeft(startingPoint - currShape.VisualLeft + spaceBetweenShapes);
                        }
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        rightMostRef = refShape.VisualLeft + refShape.AbsoluteWidth;
                        currShape.IncrementLeft(rightMostRef - currShape.VisualLeft + spaceBetweenShapes);
                    }
                }
                return;
            }

            if (isFirstShape && isObjectCenter)
            {
                if (selectedShapes.Count < 2)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanTwoSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualLeft;
                referenceWidth = shapesToDistribute[0].AbsoluteWidth;
                shapesToDistribute.RemoveAt(0);
                shapesToDistribute = ShapeUtil.SortShapesByLeft(shapesToDistribute);

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeWidth += s.AbsoluteWidth;
                }

                if (totalShapeWidth > referenceWidth)
                {
                    PPShape leftMostShape = shapesToDistribute[0];
                    PPShape rightMostShape = shapesToDistribute[shapesToDistribute.Count - 1];

                    leftMostShape.IncrementLeft(startingPoint - leftMostShape.VisualLeft);
                    rightMostShape.IncrementLeft(startingPoint + referenceWidth - rightMostShape.VisualLeft - rightMostShape.AbsoluteWidth);

                    referenceWidth = referenceWidth - (leftMostShape.AbsoluteWidth / 2) - (rightMostShape.AbsoluteWidth / 2);

                    spaceBetweenShapes = referenceWidth / (shapesToDistribute.Count - 1);
                    startingPoint = leftMostShape.VisualCenter.X;
                    shapesToDistribute.RemoveAt(shapesToDistribute.Count - 1);
                    shapesToDistribute.RemoveAt(0);
                }
                else
                {
                    spaceBetweenShapes = referenceWidth / (shapesToDistribute.Count + 1);
                }

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];
                    if (i == 0)
                    {
                        currShape.IncrementLeft(startingPoint - currShape.VisualCenter.X + spaceBetweenShapes);
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        rightMostRef = refShape.ActualCenter.X;
                        currShape.IncrementLeft(rightMostRef - currShape.VisualCenter.X + spaceBetweenShapes);
                    }
                }
                return;
            }

            if (isFirstTwoShapes && isObjectBoundary)
            {
                if (selectedShapes.Count < 3)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanThreeSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);

                if (shapesToDistribute[0].VisualLeft > shapesToDistribute[1].VisualLeft)
                {
                    PPShape temp = shapesToDistribute[0];
                    shapesToDistribute[0] = shapesToDistribute[1];
                    shapesToDistribute[1] = temp;
                }

                startingPoint = shapesToDistribute[0].VisualLeft + shapesToDistribute[0].AbsoluteWidth;
                referenceWidth = shapesToDistribute[1].VisualLeft + shapesToDistribute[1].AbsoluteWidth - shapesToDistribute[0].VisualLeft;

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeWidth += s.AbsoluteWidth;
                }
                spaceBetweenShapes = (referenceWidth - totalShapeWidth) / (shapesToDistribute.Count - 1);

                shapesToDistribute.RemoveAt(1);
                shapesToDistribute.RemoveAt(0);
                shapesToDistribute = ShapeUtil.SortShapesByLeft(shapesToDistribute);

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];
                    if (i==0)
                    {
                        currShape.IncrementLeft(startingPoint - currShape.VisualLeft + spaceBetweenShapes);
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        rightMostRef = refShape.VisualLeft + refShape.AbsoluteWidth;
                        currShape.IncrementLeft(rightMostRef - currShape.VisualLeft + spaceBetweenShapes);
                    }
                }
                return;
            }

            if (isFirstTwoShapes && isObjectCenter)
            {
                if (selectedShapes.Count < 3)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanThreeSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);

                if (shapesToDistribute[0].VisualLeft > shapesToDistribute[1].VisualLeft)
                {
                    PPShape temp = shapesToDistribute[0];
                    shapesToDistribute[0] = shapesToDistribute[1];
                    shapesToDistribute[1] = temp;
                }

                startingPoint = shapesToDistribute[0].VisualCenter.X;
                referenceWidth = shapesToDistribute[1].VisualCenter.X -shapesToDistribute[0].VisualCenter.X;

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeWidth += s.AbsoluteWidth;
                }
                
                spaceBetweenShapes = referenceWidth / (shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(1);
                shapesToDistribute.RemoveAt(0);

                shapesToDistribute = ShapeUtil.SortShapesByLeft(shapesToDistribute);

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];
                    if (i == 0)
                    {
                        currShape.IncrementLeft(startingPoint - currShape.VisualCenter.X + spaceBetweenShapes);
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        rightMostRef = refShape.ActualCenter.X;
                        currShape.IncrementLeft(rightMostRef - currShape.VisualCenter.X + spaceBetweenShapes);
                    }
                }
                return;
            }

            if (isExtremeShape && isObjectBoundary)
            {
                if (selectedShapes.Count < 3)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanThreeSelection);
                }

                shapesToDistribute = ShapeUtil.SortShapesByLeft(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualLeft + shapesToDistribute[0].AbsoluteWidth;
                PPShape rightMostShape = shapesToDistribute[shapesToDistribute.Count - 1];
                referenceWidth = rightMostShape.VisualLeft + rightMostShape.AbsoluteWidth - shapesToDistribute[0].VisualLeft;

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeWidth += s.AbsoluteWidth;
                }
                spaceBetweenShapes = (referenceWidth - totalShapeWidth) / (shapesToDistribute.Count - 1);

                shapesToDistribute.RemoveAt(shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(0);

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];
                    if (i == 0)
                    {
                        currShape.IncrementLeft(startingPoint - currShape.VisualLeft + spaceBetweenShapes);
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        rightMostRef = refShape.VisualLeft + refShape.AbsoluteWidth;
                        currShape.IncrementLeft(rightMostRef - currShape.VisualLeft + spaceBetweenShapes);
                    }
                }
                return;
            }

            if (isExtremeShape && isObjectCenter)
            {
                if (selectedShapes.Count < 3)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanThreeSelection);
                }

                shapesToDistribute = ShapeUtil.SortShapesByLeft(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualCenter.X;
                referenceWidth = shapesToDistribute[shapesToDistribute.Count-1].VisualCenter.X - shapesToDistribute[0].VisualCenter.X;

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeWidth += s.AbsoluteWidth;
                }

                spaceBetweenShapes = referenceWidth / (shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(0);

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];
                    if (i == 0)
                    {
                        currShape.IncrementLeft(startingPoint - currShape.VisualCenter.X + spaceBetweenShapes);
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        rightMostRef = refShape.ActualCenter.X;
                        currShape.IncrementLeft(rightMostRef - currShape.VisualCenter.X + spaceBetweenShapes);
                    }
                }
                return;
            }
        }

        public static void DistributeVertical(List<PPShape> selectedShapes, float slideHeight)
        {
            bool isSlide = PositionsLabSettings.DistributeReference == PositionsLabSettings.DistributeReferenceObject.Slide;
            bool isFirstShape = PositionsLabSettings.DistributeReference == PositionsLabSettings.DistributeReferenceObject.FirstShape;
            bool isExtremeShape = PositionsLabSettings.DistributeReference == PositionsLabSettings.DistributeReferenceObject.ExtremeShapes;
            bool isFirstTwoShapes = PositionsLabSettings.DistributeReference == PositionsLabSettings.DistributeReferenceObject.FirstTwoShapes;
            bool isObjectCenter = PositionsLabSettings.DistributeSpaceReference == PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            bool isObjectBoundary = PositionsLabSettings.DistributeSpaceReference == PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary;

            List<PPShape> shapesToDistribute;
            PPShape refShape;
            float referenceHeight, spaceBetweenShapes, startingPoint, bottomMostRef, totalShapeHeight = 0;

            if (isSlide && isObjectBoundary)
            {
                if (selectedShapes.Count < 1)
                {
                    throw new Exception(PositionsLabText.ErrorNoSelection);
                }

                startingPoint = 0;
                referenceHeight = slideHeight;
                shapesToDistribute = ShapeUtil.SortShapesByTop(selectedShapes);
                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeHeight += s.AbsoluteHeight;
                }

                // Check if need leading and trailing space
                if (totalShapeHeight > referenceHeight)
                {
                    spaceBetweenShapes = (referenceHeight - totalShapeHeight) / (shapesToDistribute.Count - 1);
                }
                else
                {
                    spaceBetweenShapes = (referenceHeight - totalShapeHeight) / (shapesToDistribute.Count + 1);
                }

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];

                    // Check if need leading and trailing space
                    if (i == 0)
                    {
                        if (spaceBetweenShapes < 0)
                        {
                            currShape.IncrementTop(startingPoint - currShape.VisualTop);
                        }
                        else
                        {
                            currShape.IncrementTop(startingPoint - currShape.VisualTop + spaceBetweenShapes);
                        }
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        bottomMostRef = refShape.VisualTop + refShape.AbsoluteHeight;
                        currShape.IncrementTop(bottomMostRef - currShape.VisualTop + spaceBetweenShapes);
                    }
                }
                return;
            }

            if (isSlide && isObjectCenter)
            {
                if (selectedShapes.Count < 1)
                {
                    throw new Exception(PositionsLabText.ErrorNoSelection);
                }

                startingPoint = 0;
                referenceHeight = slideHeight;
                shapesToDistribute = ShapeUtil.SortShapesByTop(selectedShapes);

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeHeight += s.AbsoluteHeight;
                }

                if (totalShapeHeight > referenceHeight)
                {
                    PPShape topMostShape = shapesToDistribute[0];
                    PPShape bottomMostShape = shapesToDistribute[shapesToDistribute.Count - 1];

                    topMostShape.IncrementTop(startingPoint - topMostShape.VisualTop);
                    bottomMostShape.IncrementTop(referenceHeight - bottomMostShape.VisualTop - bottomMostShape.AbsoluteHeight);

                    referenceHeight = referenceHeight - (topMostShape.AbsoluteHeight / 2)- (bottomMostShape.AbsoluteHeight / 2);

                    spaceBetweenShapes = referenceHeight / (shapesToDistribute.Count - 1);
                    startingPoint = topMostShape.VisualCenter.Y;
                    shapesToDistribute.RemoveAt(shapesToDistribute.Count - 1);
                    shapesToDistribute.RemoveAt(0);
                }
                else
                {
                    spaceBetweenShapes = referenceHeight / (shapesToDistribute.Count + 1);
                }

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];
                    if (i == 0)
                    {
                        currShape.IncrementTop(startingPoint - currShape.VisualCenter.Y + spaceBetweenShapes);
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        bottomMostRef = refShape.ActualCenter.Y;
                        currShape.IncrementTop(bottomMostRef - currShape.VisualCenter.Y + spaceBetweenShapes);
                    }
                }
                return;
            }

            if (isFirstShape && isObjectBoundary)
            {
                if (selectedShapes.Count < 2)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanTwoSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualTop;
                referenceHeight = shapesToDistribute[0].AbsoluteHeight;
                shapesToDistribute.RemoveAt(0);
                shapesToDistribute = ShapeUtil.SortShapesByTop(shapesToDistribute);

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeHeight += s.AbsoluteHeight;
                }

                // Check if need leading and trailing space
                if (totalShapeHeight > referenceHeight)
                {
                    spaceBetweenShapes = (referenceHeight - totalShapeHeight) / (shapesToDistribute.Count - 1);
                }
                else
                {
                    spaceBetweenShapes = (referenceHeight - totalShapeHeight) / (shapesToDistribute.Count + 1);
                }

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];

                    // Check if need leading and trailing space
                    if (i == 0)
                    {
                        if (spaceBetweenShapes < 0)
                        {
                            currShape.IncrementTop(startingPoint - currShape.VisualTop);
                        }
                        else
                        {
                            currShape.IncrementTop(startingPoint - currShape.VisualTop + spaceBetweenShapes);
                        }
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        bottomMostRef = refShape.VisualTop + refShape.AbsoluteHeight;
                        currShape.IncrementTop(bottomMostRef - currShape.VisualTop + spaceBetweenShapes);
                    }
                }
                return;
            }

            if (isFirstShape && isObjectCenter)
            {
                if (selectedShapes.Count < 2)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanTwoSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualTop;
                referenceHeight = shapesToDistribute[0].AbsoluteHeight;
                shapesToDistribute.RemoveAt(0);
                shapesToDistribute = ShapeUtil.SortShapesByTop(shapesToDistribute);

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeHeight += s.AbsoluteHeight;
                }

                if (totalShapeHeight > referenceHeight)
                {
                    PPShape topMostShape = shapesToDistribute[0];
                    PPShape bottomMostShape = shapesToDistribute[shapesToDistribute.Count - 1];

                    topMostShape.IncrementTop(startingPoint - topMostShape.VisualTop);
                    bottomMostShape.IncrementTop(startingPoint + referenceHeight - bottomMostShape.VisualTop - bottomMostShape.AbsoluteHeight);

                    referenceHeight = referenceHeight - (topMostShape.AbsoluteHeight / 2) - (bottomMostShape.AbsoluteHeight / 2);

                    spaceBetweenShapes = referenceHeight / (shapesToDistribute.Count - 1);
                    startingPoint = topMostShape.VisualCenter.Y;
                    shapesToDistribute.RemoveAt(shapesToDistribute.Count - 1);
                    shapesToDistribute.RemoveAt(0);
                }
                else
                {
                    spaceBetweenShapes = referenceHeight / (shapesToDistribute.Count + 1);
                }

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];
                    if (i == 0)
                    {
                        currShape.IncrementTop(startingPoint - currShape.VisualCenter.Y + spaceBetweenShapes);
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        bottomMostRef = refShape.ActualCenter.Y;
                        currShape.IncrementTop(bottomMostRef - currShape.VisualCenter.Y + spaceBetweenShapes);
                    }
                }
                return;
            }

            if (isFirstTwoShapes && isObjectBoundary)
            {
                if (selectedShapes.Count < 3)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanThreeSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);

                if (shapesToDistribute[0].VisualTop > shapesToDistribute[1].VisualTop)
                {
                    PPShape temp = shapesToDistribute[0];
                    shapesToDistribute[0] = shapesToDistribute[1];
                    shapesToDistribute[1] = temp;
                }

                startingPoint = shapesToDistribute[0].VisualTop + shapesToDistribute[0].AbsoluteHeight;
                referenceHeight = shapesToDistribute[1].VisualTop + shapesToDistribute[1].AbsoluteHeight - shapesToDistribute[0].VisualTop;

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeHeight += s.AbsoluteHeight;
                }
                spaceBetweenShapes = (referenceHeight - totalShapeHeight) / (shapesToDistribute.Count - 1);

                shapesToDistribute.RemoveAt(1);
                shapesToDistribute.RemoveAt(0);
                shapesToDistribute = ShapeUtil.SortShapesByTop(shapesToDistribute);

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];
                    if (i == 0)
                    {
                        currShape.IncrementTop(startingPoint - currShape.VisualTop + spaceBetweenShapes);
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        bottomMostRef = refShape.VisualTop + refShape.AbsoluteHeight;
                        currShape.IncrementTop(bottomMostRef - currShape.VisualTop + spaceBetweenShapes);
                    }
                }
                return;
            }

            if (isFirstTwoShapes && isObjectCenter)
            {
                if (selectedShapes.Count < 3)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanThreeSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);

                if (shapesToDistribute[0].VisualTop > shapesToDistribute[1].VisualTop)
                {
                    PPShape temp = shapesToDistribute[0];
                    shapesToDistribute[0] = shapesToDistribute[1];
                    shapesToDistribute[1] = temp;
                }

                startingPoint = shapesToDistribute[0].VisualCenter.Y;
                referenceHeight = shapesToDistribute[1].VisualCenter.Y - shapesToDistribute[0].VisualCenter.Y;

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeHeight += s.AbsoluteHeight;
                }

                spaceBetweenShapes = referenceHeight / (shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(1);
                shapesToDistribute.RemoveAt(0);

                shapesToDistribute = ShapeUtil.SortShapesByTop(shapesToDistribute);

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];
                    if (i == 0)
                    {
                        currShape.IncrementTop(startingPoint - currShape.VisualCenter.Y + spaceBetweenShapes);
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        bottomMostRef = refShape.ActualCenter.Y;
                        currShape.IncrementTop(bottomMostRef - currShape.VisualCenter.Y + spaceBetweenShapes);
                    }
                }
                return;
            }

            if (isExtremeShape && isObjectBoundary)
            {
                if (selectedShapes.Count < 3)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanThreeSelection);
                }

                shapesToDistribute = ShapeUtil.SortShapesByTop(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualTop + shapesToDistribute[0].AbsoluteHeight;
                PPShape bottomMostShape = shapesToDistribute[shapesToDistribute.Count - 1];
                referenceHeight = bottomMostShape.VisualTop + bottomMostShape.AbsoluteHeight - shapesToDistribute[0].VisualTop;

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeHeight += s.AbsoluteHeight;
                }
                spaceBetweenShapes = (referenceHeight - totalShapeHeight) / (shapesToDistribute.Count - 1);

                shapesToDistribute.RemoveAt(shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(0);

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];
                    if (i == 0)
                    {
                        currShape.IncrementTop(startingPoint - currShape.VisualTop + spaceBetweenShapes);
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        bottomMostRef = refShape.VisualTop + refShape.AbsoluteHeight;
                        currShape.IncrementTop(bottomMostRef - currShape.VisualTop + spaceBetweenShapes);
                    }
                }
                return;
            }

            if (isExtremeShape && isObjectCenter)
            {
                if (selectedShapes.Count < 3)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanThreeSelection);
                }

                shapesToDistribute = ShapeUtil.SortShapesByTop(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualCenter.Y;
                referenceHeight = shapesToDistribute[shapesToDistribute.Count - 1].VisualCenter.Y - shapesToDistribute[0].VisualCenter.Y;

                foreach (PPShape s in shapesToDistribute)
                {
                    totalShapeHeight += s.AbsoluteHeight;
                }

                spaceBetweenShapes = referenceHeight / (shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(0);

                for (int i = 0; i < shapesToDistribute.Count; i++)
                {
                    PPShape currShape = shapesToDistribute[i];
                    if (i == 0)
                    {
                        currShape.IncrementTop(startingPoint - currShape.VisualCenter.Y + spaceBetweenShapes);
                    }
                    else
                    {
                        refShape = shapesToDistribute[i - 1];
                        bottomMostRef = refShape.ActualCenter.Y;
                        currShape.IncrementTop(bottomMostRef - currShape.VisualCenter.Y + spaceBetweenShapes);
                    }
                }
                return;
            }
        }

        public static void DistributeCenter(List<PPShape> selectedShapes, float slideWidth, float slideHeight)
        {
            DistributeHorizontal(selectedShapes, slideWidth);
            DistributeVertical(selectedShapes, slideHeight);
        }

        public static void DistributeGrid(List<PPShape> selectedShapes, int numOfRows, int numOfCols)
        {
            if (PositionsLabSettings.DistributeReference == PositionsLabSettings.DistributeReferenceObject.ExtremeShapes)
            {
                throw new Exception(PositionsLabText.ErrorFunctionNotSupportedForWithinShapes);
            }

            if (PositionsLabSettings.DistributeReference == PositionsLabSettings.DistributeReferenceObject.Slide)
            {
                throw new Exception(PositionsLabText.ErrorFunctionNotSupportedForSlide);
            }

            int rowLength = numOfCols;
            int colLength = numOfRows;

            bool isFirstTwoShapes = PositionsLabSettings.DistributeReference == PositionsLabSettings.DistributeReferenceObject.FirstTwoShapes;
            bool isFirstShape = PositionsLabSettings.DistributeReference == PositionsLabSettings.DistributeReferenceObject.FirstShape;
            bool isObjectCenter = PositionsLabSettings.DistributeSpaceReference == PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            bool isObjectBoundary = PositionsLabSettings.DistributeSpaceReference == PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary;

            int colLengthGivenFullRows = (int) Math.Ceiling((double) selectedShapes.Count/rowLength);
            if (colLength <= colLengthGivenFullRows)
            {
                if (isFirstTwoShapes && isObjectCenter)
                {
                    Drawing.PointF startAnchorCenter = selectedShapes[0].VisualCenter;
                    Drawing.PointF endAnchorCenter = selectedShapes[1].VisualCenter;
                    float rowWidth = endAnchorCenter.X - startAnchorCenter.X;
                    float colHeight = endAnchorCenter.Y - startAnchorCenter.Y;
                    DistributeGridByRowWithAnchors(selectedShapes, rowLength, colLength, rowWidth, colHeight);
                }
                else if (isFirstTwoShapes && isObjectBoundary)
                {
                    PPShape startAnchor = selectedShapes[0];
                    PPShape endAnchor = selectedShapes[1];
                    float rowWidth = endAnchor.VisualLeft + endAnchor.AbsoluteWidth - startAnchor.VisualLeft;
                    float colHeight = endAnchor.VisualTop + endAnchor.AbsoluteHeight - startAnchor.VisualTop;
                    DistributeGridByRowWithAnchorsByEdge(selectedShapes, rowLength, colLength, rowWidth, colHeight);
                }
                else if (isFirstShape && isObjectCenter)
                {
                    DistributeGridByRow(selectedShapes, rowLength, colLength);
                }
                else if (isFirstShape && isObjectBoundary)
                {
                    DistributeGridByRowByEdge(selectedShapes, rowLength, colLength);
                }
            }
            else
            {
                if (isFirstTwoShapes && isObjectCenter)
                {
                    Drawing.PointF startAnchorCenter = selectedShapes[0].VisualCenter;
                    Drawing.PointF endAnchorCenter = selectedShapes[1].VisualCenter;
                    float rowWidth = endAnchorCenter.X - startAnchorCenter.X;
                    float colHeight = endAnchorCenter.Y - startAnchorCenter.Y;
                    DistributeGridByColWithAnchors(selectedShapes, rowLength, colLength, rowWidth, colHeight);
                }
                else if (isFirstTwoShapes && isObjectBoundary)
                {
                    PPShape startAnchor = selectedShapes[0];
                    PPShape endAnchor = selectedShapes[1];
                    float rowWidth = endAnchor.VisualLeft + endAnchor.AbsoluteWidth - startAnchor.VisualLeft;
                    float colHeight = endAnchor.VisualTop + endAnchor.AbsoluteHeight - startAnchor.VisualTop;
                    DistributeGridByColWithAnchorsByEdge(selectedShapes, rowLength, colLength, rowWidth, colHeight);
                }
                else if (isFirstShape && isObjectCenter)
                {
                    DistributeGridByCol(selectedShapes, rowLength, colLength);
                }
                else if (isFirstShape && isObjectBoundary)
                {
                    DistributeGridByColByEdge(selectedShapes, rowLength, colLength);
                }
            }
        }

        public static void DistributeGridByRow(List<PPShape> selectedShapes, int rowLength, int colLength)
        {
            Drawing.PointF refPoint = selectedShapes[0].VisualCenter;

            int numShapes = selectedShapes.Count;

            int numIndicesToSkip = IndicesToSkip(numShapes, rowLength, PositionsLabSettings.DistributeGridAlignment);

            float[] rowDifferences = GetLongestWidthsOfRowsByRow(selectedShapes, rowLength, numIndicesToSkip);
            float[] colDifferences = GetLongestHeightsOfColsByRow(selectedShapes, rowLength, colLength);

            float posX = refPoint.X;
            float posY = refPoint.Y;
            int remainder = numShapes%rowLength;
            int differenceIndex = 0;

            for (int i = 0; i < numShapes; i++)
            {
                //Start of new row
                if (i%rowLength == 0 && i != 0)
                {
                    posX = refPoint.X;
                    differenceIndex = 0;
                    posY += GetSpaceBetweenShapes(i/rowLength - 1, i/rowLength, colDifferences, 
                                                    PositionsLabSettings.GridMarginTop, 
                                                    PositionsLabSettings.GridMarginBottom);
                }

                //If last row, offset by num of indices to skip
                if (numShapes - i == remainder)
                {
                    differenceIndex = numIndicesToSkip;
                    posX += GetSpaceBetweenShapes(0, differenceIndex, rowDifferences, 
                                                    PositionsLabSettings.GridMarginLeft, 
                                                    PositionsLabSettings.GridMarginRight);
                }

                PPShape currentShape = selectedShapes[i];
                currentShape.IncrementLeft(posX - currentShape.VisualCenter.X);
                currentShape.IncrementTop(posY - currentShape.VisualCenter.Y);

                posX += GetSpaceBetweenShapes(differenceIndex, differenceIndex + 1, rowDifferences, 
                                                    PositionsLabSettings.GridMarginLeft, 
                                                    PositionsLabSettings.GridMarginRight);
                differenceIndex++;
            }
        }

        public static void DistributeGridByCol(List<PPShape> selectedShapes, int rowLength, int colLength)
        {
            Drawing.PointF refPoint = selectedShapes[0].VisualCenter;

            int numShapes = selectedShapes.Count;

            int numIndicesToSkip = IndicesToSkip(numShapes, colLength, PositionsLabSettings.DistributeGridAlignment);

            float[] rowDifferences = GetLongestWidthsOfRowsByCol(selectedShapes, rowLength, colLength, numIndicesToSkip);
            float[] colDifferences = GetLongestHeightsOfColsByCol(selectedShapes, rowLength, colLength, numIndicesToSkip);

            float posX = refPoint.X;
            float posY = refPoint.Y;
            int remainder = colLength - (rowLength*colLength - numShapes);
            int augmentedShapeIndex = 0;

            for (int i = 0; i < numShapes; i++)
            {
                //If last index and need to skip, skip index 
                if (numIndicesToSkip > 0 && IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    numIndicesToSkip--;
                    augmentedShapeIndex++;
                }

                //If last index and no more remainder, skip the rest
                if (IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    if (remainder <= 0)
                    {
                        augmentedShapeIndex++;
                    }
                    else
                    {
                        remainder--;
                    }
                }

                if (IsFirstIndexOfRow(augmentedShapeIndex, rowLength) && augmentedShapeIndex != 0)
                {
                    posX = refPoint.X;
                    posY += GetSpaceBetweenShapes(augmentedShapeIndex/rowLength - 1, augmentedShapeIndex/rowLength, colDifferences, 
                                                PositionsLabSettings.GridMarginTop, 
                                                PositionsLabSettings.GridMarginBottom);
                }

                PPShape currentShape = selectedShapes[i];
                Drawing.PointF center = currentShape.VisualCenter;
                currentShape.IncrementLeft(posX - center.X);
                currentShape.IncrementTop(posY - center.Y);

                posX += GetSpaceBetweenShapes(augmentedShapeIndex%rowLength, augmentedShapeIndex%rowLength + 1, rowDifferences,
                                                PositionsLabSettings.GridMarginLeft, 
                                                PositionsLabSettings.GridMarginRight);
                augmentedShapeIndex++;
            }
        }

        public static void DistributeGridByRowWithAnchors(List<PPShape> selectedShapes, int rowLength, int colLength, float rowWidth, float colHeight)
        {
            if (selectedShapes.Count < 2)
            {
                throw new Exception(PositionsLabText.ErrorFewerThanTwoSelection);
            }

            Drawing.PointF startingAnchor = selectedShapes[0].VisualCenter;
            Drawing.PointF endingAnchor = selectedShapes[1].VisualCenter;

            PPShape endAnchor = selectedShapes[1];
            selectedShapes.RemoveAt(1);
            selectedShapes.Add(endAnchor);

            float rowDifference = rowWidth / (rowLength - 1);
            float colDifference = colHeight / (colLength - 1);

            GridSpace[] gridSpaces = new GridSpace[selectedShapes.Count];

            for (int i = 0; i < selectedShapes.Count; i++)
            {
                gridSpaces[i] = new GridSpace(rowDifference, colDifference);
            }

            DistributeGridByRow(selectedShapes, rowLength, colLength, gridSpaces, 0, selectedShapes.Count - 1);
        }

        public static void DistributeGridByRow(List<PPShape> selectedShapes, int rowLength, int colLength, GridSpace[] gridSpaces, int start, int end)
        {
            int numShapes = selectedShapes.Count;
            int numIndicesToSkip = IndicesToSkip(numShapes, rowLength, PositionsLabSettings.DistributeGridAlignment);

            Drawing.PointF startingAnchor = selectedShapes[0].VisualCenter;

            float[] rowDifferences = GetLongestWidthsOfRowsByCol(selectedShapes, rowLength, colLength, numIndicesToSkip);
            float[] colDifferences = GetLongestHeightsOfColsByCol(selectedShapes, rowLength, colLength, numIndicesToSkip);

            float posX = startingAnchor.X;
            float posY = startingAnchor.Y;
            int remainder = numShapes % rowLength;

            for (int i = start; i < end; i++)
            {
                //Start of new row
                if (i % rowLength == 0 && i != 0)
                {
                    posX = startingAnchor.X;
                    posY += gridSpaces[i].ColDifference;
                }

                //If last row, offset by num of indices to skip
                if (numShapes - i == remainder)
                {
                    posX += numIndicesToSkip * gridSpaces[i].RowDifference;
                }

                PPShape currentShape = selectedShapes[i];
                currentShape.IncrementLeft(posX - currentShape.VisualCenter.X);
                currentShape.IncrementTop(posY - currentShape.VisualCenter.Y);

                posX += gridSpaces[i].RowDifference;
            }
        }

        public static void DistributeGridByRowWithAnchorsByEdge(List<PPShape> selectedShapes, int rowLength, int colLength, float rowWidth, float colHeight)
        {
            if (selectedShapes.Count < 2)
            {
                throw new Exception(PositionsLabText.ErrorFewerThanTwoSelection);
            }

            int numShapes = selectedShapes.Count;

            PPShape startAnchor = selectedShapes[0];
            PPShape endAnchor = selectedShapes[1];
            selectedShapes.RemoveAt(1);
            selectedShapes.Add(endAnchor);

            Drawing.PointF startingAnchor = selectedShapes[0].VisualCenter;

            float longestRow = rowWidth;
            float longestCol = colHeight;

            float[] colDifferences = GetLongestHeightsOfColsByRow(selectedShapes, rowLength, colLength);

            for (int i = 0; i < colDifferences.Length; i++)
            {
                longestCol -= colDifferences[i];
            }

            float posX = startingAnchor.X;
            float posY = startAnchor.VisualTop + colDifferences[0] / 2;
            float rowDifference = longestRow;
            float colDifference = longestCol / (colDifferences.Length - 1);

            for (int i = 0; i < numShapes - 1; i++)
            {
                //Start of new row
                if (i % rowLength == 0)
                {
                    rowDifference = longestRow;
                    int j = 0;
                    for (j = 0; j < rowLength; j++)
                    {
                        if (i + j >= numShapes)
                        {
                            break;
                        }
                        rowDifference -= selectedShapes[i + j].AbsoluteWidth;
                    }
                    if (j > 1)
                    {
                        rowDifference /= (j - 1);
                    }
                    
                    if (i != 0)
                    {
                        posX = selectedShapes[0].VisualLeft + selectedShapes[i].AbsoluteWidth / 2;
                        posY += (colDifferences[i / rowLength - 1] / 2 + colDifferences[i / rowLength] / 2 + colDifference);
                    }
                }

                PPShape currentShape = selectedShapes[i];
                currentShape.IncrementLeft(posX - currentShape.VisualCenter.X);

                if (i / rowLength == 0)
                {
                    currentShape.VisualTop = startAnchor.VisualTop;
                }
                else
                {
                    currentShape.IncrementTop(posY - currentShape.VisualCenter.Y);
                }

                if (i + 1 < numShapes)
                {
                    posX += (selectedShapes[i].AbsoluteWidth / 2 + selectedShapes[i + 1].AbsoluteWidth / 2 + rowDifference);
                }
            }
        }

        public static void DistributeGridByRowByEdge(List<PPShape> selectedShapes, int rowLength, int colLength)
        {
            int numShapes = selectedShapes.Count;

            Drawing.PointF startingAnchor = selectedShapes[0].VisualCenter;

            float posX = startingAnchor.X;
            float posY = startingAnchor.Y;

            float longestRow = GetLongestRowWidthByRow(selectedShapes, rowLength);
            float[] colDifferences = GetLongestHeightsOfColsByRow(selectedShapes, rowLength, colLength);

            float rowDifference = longestRow;

            for (int i = 0; i < numShapes; i++)
            {
                //Start of new row
                if (i % rowLength == 0)
                {
                    rowDifference = longestRow;
                    int j = 0;
                    for (j = 0; j < rowLength; j++)
                    {
                        if (i + j >= numShapes)
                        {
                            break;
                        }
                        rowDifference -= selectedShapes[i + j].AbsoluteWidth;
                    }
                    if (j > 1)
                    {
                        rowDifference /= (j - 1);
                    }
                    if (i != 0)
                    {   
                        posX = selectedShapes[0].VisualLeft + selectedShapes[i].AbsoluteWidth / 2;
                        posY += GetSpaceBetweenShapes(i / rowLength - 1, i / rowLength, colDifferences,
                                                        PositionsLabSettings.GridMarginTop,
                                                        PositionsLabSettings.GridMarginBottom);
                    }
                }

                PPShape currentShape = selectedShapes[i];
                currentShape.IncrementLeft(posX - currentShape.VisualCenter.X);
                currentShape.IncrementTop(posY - currentShape.VisualCenter.Y);

                if (i + 1 < numShapes)
                {
                    posX += (selectedShapes[i].AbsoluteWidth / 2 + selectedShapes[i + 1].AbsoluteWidth / 2 + rowDifference);
                }
            }
        }

        public static void DistributeGridByColWithAnchors(List<PPShape> selectedShapes, int rowLength, int colLength, float rowWidth, float colHeight)
        {
            if (selectedShapes.Count < 2)
            {
                throw new Exception(PositionsLabText.ErrorFewerThanTwoSelection);
            }

            Drawing.PointF startingAnchor = selectedShapes[0].VisualCenter;
            Drawing.PointF endingAnchor = selectedShapes[1].VisualCenter;

            float rowDifference = rowWidth / (rowLength - 1);
            float colDifference = colHeight / (colLength - 1);

            PPShape endAnchor = selectedShapes[1];
            selectedShapes.RemoveAt(1);
            selectedShapes.Add(endAnchor);

            GridSpace[] gridSpaces = new GridSpace[selectedShapes.Count];

            for (int i = 0; i < selectedShapes.Count; i++)
            {
                gridSpaces[i] = new GridSpace(rowDifference, colDifference);
            }

            DistributeGridByCol(selectedShapes, rowLength, colLength, gridSpaces, 0, selectedShapes.Count - 1);
        }

        public static void DistributeGridByCol(List<PPShape> selectedShapes, int rowLength, int colLength, GridSpace[] gridSpaces, int start, int end)
        {
            int numShapes = selectedShapes.Count;

            int numIndicesToSkip = IndicesToSkip(numShapes, colLength, PositionsLabSettings.DistributeGridAlignment);

            Drawing.PointF startingAnchor = selectedShapes[0].VisualCenter;

            float posX = startingAnchor.X;
            float posY = startingAnchor.Y;
            int remainder = colLength - (rowLength * colLength - numShapes);
            int augmentedShapeIndex = 0;

            for (int i = start; i < end; i++)
            {
                //If last index and need to skip, skip index 
                if (numIndicesToSkip > 0 && IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    numIndicesToSkip--;
                    augmentedShapeIndex++;
                }

                //If last index and no more remainder, skip the rest
                if (IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    if (remainder <= 0)
                    {
                        augmentedShapeIndex++;
                    }
                    else
                    {
                        remainder--;
                    }
                }

                if (IsFirstIndexOfRow(augmentedShapeIndex, rowLength) && augmentedShapeIndex != 0)
                {
                    posX = startingAnchor.X;
                    posY += gridSpaces[i].ColDifference;
                }

                PPShape currentShape = selectedShapes[i];
                Drawing.PointF center = currentShape.VisualCenter;
                currentShape.IncrementLeft(posX - center.X);
                currentShape.IncrementTop(posY - center.Y);

                posX += gridSpaces[i].RowDifference;
                augmentedShapeIndex++;
            }
        }

        public static void DistributeGridByColByEdge(List<PPShape> selectedShapes, int rowLength, int colLength)
        {
            int numShapes = selectedShapes.Count;
            Drawing.PointF startingAnchor = selectedShapes[0].VisualCenter;

            float posX = startingAnchor.X;
            float posY = startingAnchor.Y;
            int remainder = colLength - (rowLength * colLength - numShapes);
            int augmentedShapeIndex = 0;

            float longestRow = GetLongestRowWidthByCol(selectedShapes, rowLength, colLength);
            float[] colDifferences = GetLongestHeightsOfColsByCol(selectedShapes, rowLength, colLength, 0);
            float rowDifference = longestRow;

            for (int i = 0; i < numShapes; i++)
            {
                //If last index and no more remainder, skip the rest
                if (IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    if (remainder <= 0)
                    {
                        augmentedShapeIndex++;
                    }
                    else
                    {
                        remainder--;
                    }
                }

                if (IsFirstIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    rowDifference = longestRow;
                    int j = 0;
                    int end = rowLength;

                    if (remainder <= 0)
                    {
                        end--;
                    }

                    for (j = 0; j < end; j++)
                    {
                        if (i + j >= numShapes)
                        {
                            break;
                        }
                        rowDifference -= selectedShapes[i + j].AbsoluteWidth;
                    }
                    if (j > 1)
                    {
                        rowDifference /= (j - 1);
                    }

                    if (augmentedShapeIndex != 0)
                    {
                        posX = selectedShapes[0].VisualLeft + selectedShapes[i].AbsoluteWidth / 2;
                        posY += GetSpaceBetweenShapes(augmentedShapeIndex / rowLength - 1, augmentedShapeIndex / rowLength, colDifferences,
                                                        PositionsLabSettings.GridMarginTop,
                                                        PositionsLabSettings.GridMarginBottom);
                    }
                }

                PPShape currentShape = selectedShapes[i];
                Drawing.PointF center = currentShape.VisualCenter;
                currentShape.IncrementLeft(posX - center.X);
                currentShape.IncrementTop(posY - center.Y);

                if (i + 1 < numShapes)
                {
                    posX += (selectedShapes[i].AbsoluteWidth / 2 + selectedShapes[i + 1].AbsoluteWidth / 2 + rowDifference);
                }
                augmentedShapeIndex++;
            }
        }

        public static void DistributeGridByColWithAnchorsByEdge(List<PPShape> selectedShapes, int rowLength, int colLength, float rowWidth, float colHeight)
        {
            if (selectedShapes.Count < 2)
            {
                throw new Exception(PositionsLabText.ErrorFewerThanTwoSelection);
            }

            int numShapes = selectedShapes.Count;

            PPShape startAnchor = selectedShapes[0];
            PPShape endAnchor = selectedShapes[1];
            selectedShapes.RemoveAt(1);
            selectedShapes.Add(endAnchor);

            Drawing.PointF startingAnchor = selectedShapes[0].VisualCenter;

            float longestRow = rowWidth;
            float longestCol = colHeight;

            float[] colDifferences = GetLongestHeightsOfColsByCol(selectedShapes, rowLength, colLength, 0);

            for (int i = 0; i < colDifferences.Length; i++)
            {
                longestCol -= colDifferences[i];
            }

            float posX = startingAnchor.X;
            float posY = startAnchor.VisualTop + colDifferences[0] / 2;
            float rowDifference = longestRow;
            float colDifference = longestCol / (colDifferences.Length - 1);
            int remainder = colLength - (rowLength * colLength - numShapes);
            int augmentedShapeIndex = 0;

            for (int i = 0; i < numShapes - 1; i++)
            {
                //If last index and no more remainder, skip the rest
                if (IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    if (remainder <= 0)
                    {
                        augmentedShapeIndex++;
                    }
                    else
                    {
                        remainder--;
                    }
                }

                if (IsFirstIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    rowDifference = longestRow;
                    int j = 0;
                    int end = rowLength;

                    if (remainder <= 0)
                    {
                        end--;
                    }

                    for (j = 0; j < end; j++)
                    {
                        if (i + j >= numShapes)
                        {
                            break;
                        }
                        rowDifference -= selectedShapes[i + j].AbsoluteWidth;
                    }
                    if (j > 1)
                    {
                        rowDifference /= (j - 1);
                    }

                    if (augmentedShapeIndex != 0)
                    {
                        posX = selectedShapes[0].VisualLeft + selectedShapes[i].AbsoluteWidth / 2;
                        posY += (colDifferences[augmentedShapeIndex / rowLength - 1] / 2 + colDifferences[augmentedShapeIndex / rowLength] / 2 + colDifference);
                    }
                }

                PPShape currentShape = selectedShapes[i];
                Drawing.PointF center = currentShape.VisualCenter;
                currentShape.IncrementLeft(posX - center.X);

                if (augmentedShapeIndex / rowLength == 0)
                {
                    currentShape.VisualTop = startAnchor.VisualTop;
                }
                else
                {
                    currentShape.IncrementTop(posY - center.Y);
                }

                if (i + 1 < numShapes)
                {
                    posX += (selectedShapes[i].AbsoluteWidth / 2 + selectedShapes[i + 1].AbsoluteWidth / 2 + rowDifference);
                }
                augmentedShapeIndex++;
            }
        }

        public static void DistributeRadial(ShapeRange selectedShapes)
        {
            bool isAtSecondShape = PositionsLabSettings.DistributeRadialReference == PositionsLabSettings.DistributeRadialReferenceObject.AtSecondShape;
            bool isSecondThirdShape = PositionsLabSettings.DistributeRadialReference == PositionsLabSettings.DistributeRadialReferenceObject.SecondThirdShape;
            bool isObjectBoundary = PositionsLabSettings.DistributeSpaceReference == PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary;
            bool isObjectCenter = PositionsLabSettings.DistributeSpaceReference == PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            
            Drawing.PointF origin;
            float referenceAngle, startingAngle;

            if (isAtSecondShape && isObjectBoundary)
            {
                if (selectedShapes.Count < 3)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanThreeSelection);
                }

                origin = ShapeUtil.GetCenterPoint(selectedShapes[1]);

                float[] boundaryAngles = GetShapeBoundaryAngles(origin, selectedShapes[2]);
                startingAngle = boundaryAngles[1];
                float endingAngle = boundaryAngles[0];

                if (startingAngle == 0 && boundaryAngles[1] == 360)
                {
                    throw new Exception(PositionsLabText.ErrorFunctionNotSupportedForOverlapRefShapeCenter);
                }

                referenceAngle = endingAngle - startingAngle;
                if (referenceAngle < 0)
                {
                    referenceAngle += 360;
                }

                float offset = endingAngle - startingAngle;
                if (offset > 0)
                {
                    offset -= 360;
                }

                DistributeShapesWithinAngleForBoundary(selectedShapes, origin, startingAngle, referenceAngle, 3, offset: offset);
            }
            else if (isAtSecondShape && isObjectCenter)
            {
                if (selectedShapes.Count < 3)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanThreeSelection);
                }
                
                origin = ShapeUtil.GetCenterPoint(selectedShapes[1]);
                
                startingAngle = (float)AngleBetweenTwoPoints(origin, GetVisualCenter(selectedShapes[2]));
                referenceAngle = 360;

                DistributeShapesWithinAngleForCenter(selectedShapes, origin, startingAngle, referenceAngle, 3);
            }
            else if (isSecondThirdShape && isObjectBoundary)
            {
                if (selectedShapes.Count < 4)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanFourSelection);
                }

                origin = ShapeUtil.GetCenterPoint(selectedShapes[1]);
                float[] startingShapeBoundaryAngles = GetShapeBoundaryAngles(origin, selectedShapes[2]);
                float[] endingShapeBoundaryAngles = GetShapeBoundaryAngles(origin, selectedShapes[3]);
                startingAngle = startingShapeBoundaryAngles[0];
                float endingAngle = endingShapeBoundaryAngles[1];

                if ((startingAngle == 0 && startingShapeBoundaryAngles[1] == 360)
                    || (endingShapeBoundaryAngles[0] == 0 && endingAngle == 360))
                {
                    throw new Exception(PositionsLabText.ErrorFunctionNotSupportedForOverlapRefShapeCenter);
                }

                float startingShapeAngle = startingShapeBoundaryAngles[1] - startingAngle;
                if (startingShapeAngle < 0)
                {
                    startingShapeAngle += 360;
                }

                float endingShapeAngle = endingAngle - endingShapeBoundaryAngles[0];
                if (endingShapeAngle < 0)
                {
                    endingShapeAngle += 360;
                }

                referenceAngle = endingAngle - startingAngle;
                if (referenceAngle < 0)
                {
                    referenceAngle += 360;
                }

                DistributeShapesWithinAngleForBoundary(selectedShapes, origin, startingAngle, referenceAngle, 4, startingShapeAngle,
                    endingShapeAngle);
            }
            else if (isSecondThirdShape && isObjectCenter)
            {
                if (selectedShapes.Count < 4)
                {
                    throw new Exception(PositionsLabText.ErrorFewerThanFourSelection);
                }

                origin = ShapeUtil.GetCenterPoint(selectedShapes[1]);
                startingAngle = (float)AngleBetweenTwoPoints(origin, GetVisualCenter(selectedShapes[2]));
                float endingAngle = (float)AngleBetweenTwoPoints(origin, GetVisualCenter(selectedShapes[3]));

                referenceAngle = endingAngle - startingAngle;
                if (referenceAngle < 0)
                {
                    referenceAngle += 360;
                }

                DistributeShapesWithinAngleForCenter(selectedShapes, origin, startingAngle, referenceAngle, 4);
            }
        }

        public static void DistributeShapesWithinAngleForCenter(ShapeRange selectedShapes, Drawing.PointF origin, float startingAngle,
            float referenceAngle, int startingIndex)
        {
            List<ShapeAngleInfo> shapeAngleInfos = new List<ShapeAngleInfo>();

            for (int i = startingIndex; i <= selectedShapes.Count; i++)
            {
                float angle = (float)AngleBetweenTwoPoints(origin, ShapeUtil.GetCenterPoint(selectedShapes[i]));
                float angleFromStart = (angle + (360 - startingAngle)) % 360;

                ShapeAngleInfo shapeAngleInfo = new ShapeAngleInfo(selectedShapes[i], angleFromStart);
                shapeAngleInfos.Add(shapeAngleInfo);
            }

            shapeAngleInfos = shapeAngleInfos.OrderBy(x => x.Angle).ToList();

            float angleBetweenShapes = referenceAngle / (shapeAngleInfos.Count + 1);
            float endingAngle = 0f;

            foreach (ShapeAngleInfo shapeAngleInfo in shapeAngleInfos)
            {
                endingAngle += angleBetweenShapes;

                float rotationAngle = endingAngle - shapeAngleInfo.Angle;
                Rotate(shapeAngleInfo.Shape, origin, rotationAngle, PositionsLabSettings.DistributeShapeOrientation);
            }
        }

        public static void DistributeShapesWithinAngleForBoundary(ShapeRange selectedShapes, Drawing.PointF origin, float startingAngle,
             float referenceAngle, int startingIndex, float startingShapeAngle = 0, float endingShapeAngle = 0, float offset = 0)
        {
            List<ShapeAngleInfo> shapeAngleInfos = new List<ShapeAngleInfo>();

            float boundaryShapeAngle = startingShapeAngle + endingShapeAngle;
            if (boundaryShapeAngle >= referenceAngle)
            {
                boundaryShapeAngle = 0;
            }

            int count = 0;
            bool isStable = false;
            while (!isStable && count < 20)
            {
                float totalShapeAngle = boundaryShapeAngle;

                if (count == 0)
                {
                    for (int i = startingIndex; i <= selectedShapes.Count; i++)
                    {
                        float shapeAngle;
                        float angle = GetShapeAngleInfo(selectedShapes[i], origin, startingAngle, totalShapeAngle, out shapeAngle);
                        totalShapeAngle += shapeAngle;

                        ShapeAngleInfo shapeAngleInfo = new ShapeAngleInfo(selectedShapes[i], angle, shapeAngle);
                        shapeAngleInfos.Add(shapeAngleInfo);
                    }

                    shapeAngleInfos = shapeAngleInfos.OrderBy(x => (x.Angle - offset) % 360).ToList();
                }
                else
                {
                    foreach (ShapeAngleInfo shapeAngleInfo in shapeAngleInfos)
                    {
                        float shapeAngle;
                        float angle = GetShapeAngleInfo(shapeAngleInfo.Shape, origin, startingAngle, totalShapeAngle, out shapeAngle);
                        totalShapeAngle += shapeAngle;

                        shapeAngleInfo.Angle = angle;
                        shapeAngleInfo.ShapeAngle = shapeAngle;
                    }
                }

                float angleBetweenShapes = (referenceAngle - totalShapeAngle) / (shapeAngleInfos.Count + 1);
                float endingAngle = (boundaryShapeAngle == 0) ? angleBetweenShapes : startingShapeAngle + angleBetweenShapes;

                isStable = true;

                foreach (ShapeAngleInfo shapeAngleInfo in shapeAngleInfos)
                {
                    float rotationAngle = (endingAngle - shapeAngleInfo.Angle) % 360;
                    if (rotationAngle > threshold || rotationAngle < -threshold)
                    {
                        isStable = false;
                        Rotate(shapeAngleInfo.Shape, origin, rotationAngle, PositionsLabSettings.DistributeShapeOrientation);
                    }

                    endingAngle += shapeAngleInfo.ShapeAngle + angleBetweenShapes;
                }

                count++;
            }
        }

        private static float GetShapeAngleInfo(Shape shape, Drawing.PointF origin, float startingAngle, float totalShapeAngle,
            out float shapeAngle)
        {
            float[] boundaryAnglesFromStart = GetShapeBoundaryAngles(origin, shape);
            if (boundaryAnglesFromStart[0] == 0 && boundaryAnglesFromStart[1] == 360)
            {
                throw new Exception(PositionsLabText.ErrorFunctionNotSupportedForOverlapRefShapeCenter);
            }

            boundaryAnglesFromStart[0] = (boundaryAnglesFromStart[0] + (360 - startingAngle)) % 360;
            boundaryAnglesFromStart[1] = (boundaryAnglesFromStart[1] + (360 - startingAngle)) % 360;

            shapeAngle = boundaryAnglesFromStart[1] - boundaryAnglesFromStart[0];
            if (boundaryAnglesFromStart[0] > boundaryAnglesFromStart[1])
            {
                shapeAngle += 360;
            }

            return boundaryAnglesFromStart[0];
        }

        #endregion

        #region Swap

        public static void Swap(List<PPShape> selectedShapes, bool isPreview)
        {
            if (selectedShapes.Count < 2)
            {
                throw new Exception(PositionsLabText.ErrorFewerThanTwoSelection);
            }

            List<PPShape> sortedShapes = selectedShapes;

            if (!PositionsLabSettings.IsSwapByClickOrder)
            {
                if (ListIsPreviouslySelected(selectedShapes, prevSelectedShapes))
                {
                    sortedShapes = SortPPShapesByName(selectedShapes, prevSortedShapeNames);
                }
                else
                {
                    sortedShapes = ShapeUtil.SortShapesByLeft(selectedShapes);
                }
            }
            else
            {
                prevSelectedShapes.Clear();
            }

            Drawing.PointF firstPos = GetSwapReferencePoint(sortedShapes[0], PositionsLabSettings.SwapReferencePoint);

            List<string> shapeNames = new List<string>();

            for (int i = 0; i < sortedShapes.Count; i++)
            {
                PPShape currentShape = sortedShapes[i];
                if (i < sortedShapes.Count - 1)
                {
                    Drawing.PointF currentPos = GetSwapReferencePoint(currentShape, PositionsLabSettings.SwapReferencePoint);
                    Drawing.PointF nextPos = GetSwapReferencePoint(sortedShapes[i + 1], PositionsLabSettings.SwapReferencePoint);
                    currentShape.IncrementLeft(nextPos.X - currentPos.X);
                    currentShape.IncrementTop(nextPos.Y - currentPos.Y);
                    ShapeUtil.SwapZOrder(currentShape._shape, sortedShapes[i + 1]._shape);
                }
                else
                {
                    Drawing.PointF currentPos = GetSwapReferencePoint(currentShape, PositionsLabSettings.SwapReferencePoint);
                    currentShape.IncrementLeft(firstPos.X - currentPos.X);
                    currentShape.IncrementTop(firstPos.Y - currentPos.Y);
                }

                if (i != 0 && !PositionsLabSettings.IsSwapByClickOrder && !isPreview)
                {
                    shapeNames.Add(currentShape.Name);
                }
            }

            if (!PositionsLabSettings.IsSwapByClickOrder && !isPreview)
            {
                shapeNames.Insert(0, sortedShapes[0].Name);
                prevSortedShapeNames = shapeNames;
                SaveSelectedList(selectedShapes, prevSelectedShapes);
            }
        }

        #endregion

        #region Adjustment
        
        public static void Rotate(Shape shape, Drawing.PointF origin, float angle, PositionsLabSettings.RadialShapeOrientationObject shapeOrientation)
        {
            Drawing.PointF unrotatedCenter = ShapeUtil.GetCenterPoint(shape);
            Drawing.PointF rotatedCenter = CommonUtil.RotatePoint(unrotatedCenter, origin, angle);

            shape.Left += (rotatedCenter.X - unrotatedCenter.X);
            shape.Top += (rotatedCenter.Y - unrotatedCenter.Y);

            if (shapeOrientation == PositionsLabSettings.RadialShapeOrientationObject.Dynamic)
            {
                shape.Rotation = AddAngles(shape.Rotation, angle);
            }
        }

        #endregion

        #region Snap

        public static void SnapVertical(IList<Shape> selectedShapes)
        {
            foreach (Shape s in selectedShapes)
            {
                SnapShapeVertical(s);
            }
        }

        public static void SnapHorizontal(IList<Shape> selectedShapes)
        {
            foreach (Shape s in selectedShapes)
            {
                SnapShapeHorizontal(s);
            }
        }

        public static void SnapAway(IList<Shape> shapes)
        {
            if (shapes.Count < 2)
            {
                throw new Exception(PositionsLabText.ErrorFewerThanTwoSelection);
            }

            Drawing.PointF refShapeCenter = ShapeUtil.GetCenterPoint(shapes[0]);
            bool isAllSameDir = true;
            int lastDir = -1;

            for (int i = 1; i < shapes.Count; i++)
            {
                Shape shape = shapes[i];
                Drawing.PointF shapeCenter = ShapeUtil.GetCenterPoint(shape);
                float angle = (float) AngleBetweenTwoPoints(refShapeCenter, shapeCenter);

                int dir = GetDirectionWrtRefShape(shape, angle);

                if (i == 1)
                {
                    lastDir = dir;
                }

                if (!IsSameDirection(lastDir, dir))
                {
                    isAllSameDir = false;
                    break;
                }
            }

            if (!isAllSameDir || lastDir == None || lastDir == Up)
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
                Drawing.PointF shapeCenter = ShapeUtil.GetCenterPoint(shape);
                float angle = (float) AngleBetweenTwoPoints(refShapeCenter, shapeCenter);

                float defaultUpAngle = 0;
                bool hasDefaultDirection = shapeDefaultUpAngle.TryGetValue(shape.AutoShapeType, out defaultUpAngle);

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

                shape.Rotation = (defaultUpAngle + angle) + lastDir * 90;
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

        private static void SnapTo0Or180(Shape shape)
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
            bool shapeIsVertical = shape.Height > shape.Width;

            if (NearlyEqual(shape.Height, shape.Width, Epsilon))
            {
                float defaultUpAngle = 0;
                bool hasDefaultDirection = shapeDefaultUpAngle.TryGetValue(shape.AutoShapeType, out defaultUpAngle);
                if (hasDefaultDirection)
                {
                    if (NearlyEqual(defaultUpAngle, 0.0f, Epsilon) || NearlyEqual(defaultUpAngle, 180.0f, Epsilon))
                    {
                        shapeIsVertical = true;
                    }
                    else
                    {
                        shapeIsVertical = false;
                    }
                }
            }

            return shapeIsVertical;
        }

        public static void FlipHorizontal(ShapeRange selectedShapes)
        {
            if (selectedShapes.Count < 1)
            {
                throw new Exception(PositionsLabText.ErrorNoSelection);
            }

            for (int i = 1; i <= selectedShapes.Count; i++)
            {
                Shape currentShape = selectedShapes[i];
                float currentRotation = currentShape.Rotation;
                currentShape.Flip(MsoFlipCmd.msoFlipHorizontal);
                currentShape.Rotation = currentRotation;
            }
        }

        public static void FlipVertical(ShapeRange selectedShapes)
        {
            if (selectedShapes.Count < 1)
            {
                throw new Exception(PositionsLabText.ErrorNoSelection);
            }

            for (int i = 1; i <= selectedShapes.Count; i++)
            {
                Shape currentShape = selectedShapes[i];
                float currentRotation = currentShape.Rotation;
                currentShape.Flip(MsoFlipCmd.msoFlipVertical);
                currentShape.Rotation = currentRotation;
            }
        }

        #endregion

        #endregion

        #region Util

        public static double AngleBetweenTwoPoints(Drawing.PointF refPoint, Drawing.PointF pt)
        {
            double angle = Math.Atan((pt.Y - refPoint.Y)/(pt.X - refPoint.X))*180/Math.PI;

            if (pt.X - refPoint.X >= 0)
            {
                angle = 90 + angle;
            }
            else
            {
                angle = 270 + angle;
            }

            return angle;
        }

        public static double DistanceBetweenTwoPoints(Drawing.PointF refPoint, Drawing.PointF pt)
        {
            double distance = Math.Sqrt(Math.Pow(pt.X - refPoint.X, 2) + Math.Pow(refPoint.Y - pt.Y, 2));
            return distance;
        }

        public static bool NearlyEqual(float a, float b, float epsilon)
        {
            float absA = Math.Abs(a);
            float absB = Math.Abs(b);
            float diff = Math.Abs(a - b);

            if (a == b)
            {
                // shortcut, handles infinities
                return true;
            }
            if (a == 0 || b == 0 || diff < float.Epsilon)
            {
                // a or b is zero or both are extremely close to it
                // relative error is less meaningful here
                return diff < epsilon;
            }
            // use relative error
            return diff/(absA + absB) < epsilon;
        }

        private static int GetDirectionWrtRefShape(Shape shape, float angleFromRefShape)
        {
            float defaultUpAngle;
            bool hasDefaultDirection = shapeDefaultUpAngle.TryGetValue(shape.AutoShapeType, out defaultUpAngle);

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
            float phaseInFloat = diff/90;

            if (!NearlyEqual(phaseInFloat, (float) Math.Round(phaseInFloat), Epsilon))
            {
                return None;
            }

            int phase = (int) Math.Round(phaseInFloat);

            return phase % 4;
        }

        private static bool IsSameDirection(int a, int b)
        {
            return (a == b);
        }

        public static float AddAngles(float a, float b)
        {
            return (a + b)%360;
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

        public static float[] GetLongestWidthsOfRowsByRow(List<PPShape> shapes, int rowLength, int numIndicesToSkip)
        {
            float[] longestWidths = new float[rowLength];
            int numShapes = shapes.Count;
            int remainder = numShapes%rowLength;

            for (int i = 0; i < numShapes; i++)
            {
                int longestRowIndex = i%rowLength;
                if (numShapes - i == remainder - 1)
                {
                    longestRowIndex += numIndicesToSkip;
                }
                if (longestWidths[longestRowIndex] < shapes[i].AbsoluteWidth)
                {
                    longestWidths[longestRowIndex] = shapes[i].AbsoluteWidth;
                }
            }

            return longestWidths;
        }

        public static float GetLongestRowWidthByRow(List<PPShape> shapes, int rowLength)
        {
            float longestRow = 0.0f;
            float longestRowSoFar = 0.0f;
            int numShapes = shapes.Count;
            int remainder = numShapes % rowLength;

            float gridMarginLeft = PositionsLabSettings.GridMarginLeft;
            float gridMarginRight = PositionsLabSettings.GridMarginRight;

            for (int i = 0; i < numShapes; i++)
            {
                int rowIndex = i % rowLength;

                if (rowIndex == 0)
                {
                    if (longestRowSoFar > longestRow)
                    {
                        longestRow = longestRowSoFar;
                    }
                    longestRowSoFar = -(gridMarginLeft + gridMarginRight);
                }
                longestRowSoFar += (shapes[i].AbsoluteWidth + gridMarginLeft + gridMarginRight);
            }

            if (longestRowSoFar > longestRow)
            {
                longestRow = longestRowSoFar;
            }

            return longestRow;
        }

        public static float[] GetLongestHeightsOfColsByRow(List<PPShape> shapes, int rowLength, int colLength)
        {
            float[] longestHeights = new float[colLength];

            for (int i = 0; i < shapes.Count; i++)
            {
                int longestHeightIndex = i/rowLength;
                if (longestHeights[longestHeightIndex] < shapes[i].AbsoluteHeight)
                {
                    longestHeights[longestHeightIndex] = shapes[i].AbsoluteHeight;
                }
            }

            return longestHeights;
        }

        public static float[] GetLongestWidthsOfRowsByCol(List<PPShape> shapes, int rowLength, int colLength, int numIndicesToSkip)
        {
            float[] longestWidths = new float[rowLength];
            int numShapes = shapes.Count;
            int augmentedShapeIndex = 0;
            int remainder = colLength - (rowLength*colLength - numShapes);

            for (int i = 0; i < numShapes; i++)
            {
                //If last index and need to skip, skip index 
                if (numIndicesToSkip > 0 && IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    numIndicesToSkip--;
                    augmentedShapeIndex++;
                }

                //If last index and no more remainder, skip the rest
                if (IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    if (remainder <= 0)
                    {
                        augmentedShapeIndex++;
                    }
                    else
                    {
                        remainder--;
                    }
                }

                int longestWidthsArrayIndex = augmentedShapeIndex%rowLength;

                if (longestWidths[longestWidthsArrayIndex] < shapes[i].AbsoluteWidth)
                {
                    longestWidths[longestWidthsArrayIndex] = shapes[i].AbsoluteWidth;
                }

                augmentedShapeIndex++;
            }

            return longestWidths;
        }

        public static float GetLongestRowWidthByCol(List<PPShape> shapes, int rowLength, int colLength)
        {
            float longestWidth = 0;
            int numShapes = shapes.Count;
            
            int augmentedShapeIndex = 0;
            int remainder = colLength - (rowLength * colLength - numShapes);
            float rowSoFar = 0;

            float gridMarginLeft = PositionsLabSettings.GridMarginLeft;
            float gridMarginRight = PositionsLabSettings.GridMarginRight;

            for (int i = 0; i < numShapes; i++)
            {
                //If last index and no more remainder, skip the rest
                if (IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    if (remainder <= 0)
                    {
                        augmentedShapeIndex++;
                    }
                    else
                    {
                        remainder--;
                        rowSoFar += (shapes[i].AbsoluteWidth + gridMarginLeft + gridMarginRight);
                        augmentedShapeIndex++;
                        i++;
                    }

                    if (rowSoFar > longestWidth)
                    {
                        longestWidth = rowSoFar;
                    }
                    rowSoFar = -(gridMarginLeft + gridMarginRight);
                }

                if (i < numShapes)
                {
                    rowSoFar += (shapes[i].AbsoluteWidth + gridMarginLeft + gridMarginRight);
                }                
                augmentedShapeIndex++;
            }

            if (rowSoFar > longestWidth)
            {
                longestWidth = rowSoFar;
            }

            return longestWidth;
        }

        public static float[] GetLongestHeightsOfColsByCol(List<PPShape> shapes, int rowLength, int colLength, int numIndicesToSkip)
        {
            float[] longestHeights = new float[colLength];
            int numShapes = shapes.Count;
            int augmentedShapeIndex = 0;
            int remainder = colLength - (rowLength*colLength - numShapes);

            for (int i = 0; i < numShapes; i++)
            {
                //If last index and need to skip, skip index 
                if (numIndicesToSkip > 0 && IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    numIndicesToSkip--;
                    augmentedShapeIndex++;
                }

                //If last index and no more remainder, skip the rest
                if (IsLastIndexOfRow(augmentedShapeIndex, rowLength))
                {
                    if (remainder <= 0)
                    {
                        augmentedShapeIndex++;
                    }
                    else
                    {
                        remainder--;
                    }
                }

                int longestHeightArrayIndex = augmentedShapeIndex/rowLength;

                if (longestHeights[longestHeightArrayIndex] < shapes[i].AbsoluteHeight)
                {
                    longestHeights[longestHeightArrayIndex] = shapes[i].AbsoluteHeight;
                }

                augmentedShapeIndex++;
            }

            return longestHeights;
        }

        private static bool IsFirstIndexOfRow(int index, int rowLength)
        {
            return index%rowLength == 0;
        }

        private static bool IsLastIndexOfRow(int index, int rowLength)
        {
            return index%rowLength == rowLength - 1;
        }

        public static int IndicesToSkip(int totalSelectedShapes, int rowLength, PositionsLabSettings.GridAlignment alignment)
        {
            int numOfShapesInLastRow = totalSelectedShapes%rowLength;

            if (alignment == PositionsLabSettings.GridAlignment.AlignLeft || 
                alignment == PositionsLabSettings.GridAlignment.None || 
                numOfShapesInLastRow == 0)
            {
                return 0;
            }

            if (alignment == PositionsLabSettings.GridAlignment.AlignRight)
            {
                return rowLength - numOfShapesInLastRow;
            }

            if (alignment == PositionsLabSettings.GridAlignment.AlignCenter)
            {
                int difference = rowLength - numOfShapesInLastRow;
                return difference/2;
            }

            return 0;
        }

        private static float GetSpaceBetweenShapes(int index1, int index2, float[] differences, float margin1, float margin2)
        {
            if (index1 >= differences.Length || index2 >= differences.Length)
            {
                return -1;
            }

            int start = 0;
            int end = 0;

            if (index1 < index2)
            {
                start = index1;
                end = index2;
            }
            else
            {
                start = index2;
                end = index1;
            }

            float difference = 0;

            for (int i = start; i < end; i++)
            {
                difference += (differences[i] / 2 + margin1 + margin2 + differences[i + 1] / 2);
            }

            return difference;
        }

        private static float[] GetShapeBoundaryAngles(Drawing.PointF origin, Shape shape)
        {
            PPShape ppShape = new PPShape(shape, false);
            List<Drawing.PointF> points = ppShape.Points;
            List<float> pointAngles = new List<float>();

            foreach (Drawing.PointF point in points)
            {
                float angle = (float)AngleBetweenTwoPoints(origin, point);
                pointAngles.Add(angle);
            }

            bool isSpanAcross0Degrees = false;
            bool hasTurningPoint = false;
            bool isCurrentClockwise;
            bool isPreviousClockwise = pointAngles[0] - pointAngles[pointAngles.Count - 1] >= 0;
            List<float> turningPointAngles = new List<float>();

            float[] boundaryAngles = new float[2];
            boundaryAngles[0] = pointAngles[0];
            boundaryAngles[1] = pointAngles[0];

            for (int i = 1; i < pointAngles.Count; i++)
            {
                float previousAngle = pointAngles[i - 1];
                float currentAngle = pointAngles[i];
                isCurrentClockwise = currentAngle - previousAngle >= 0;

                if (Math.Abs(currentAngle - previousAngle) > 180)
                {
                    isCurrentClockwise = !isCurrentClockwise;

                    if (isCurrentClockwise && (!isSpanAcross0Degrees || currentAngle > boundaryAngles[1]))
                    {
                        boundaryAngles[1] = currentAngle;
                    }
                    else if (!isCurrentClockwise && (!isSpanAcross0Degrees || currentAngle < boundaryAngles[0]))
                    {
                        boundaryAngles[0] = currentAngle;
                    }

                    isSpanAcross0Degrees = true;
                }

                if (isCurrentClockwise != isPreviousClockwise)
                {
                    hasTurningPoint = true;

                    if (isPreviousClockwise && previousAngle > boundaryAngles[1]
                        && (!isSpanAcross0Degrees || previousAngle < boundaryAngles[0]))
                    {
                        boundaryAngles[1] = previousAngle;
                    }

                    if (!isPreviousClockwise && previousAngle < boundaryAngles[0]
                        && (!isSpanAcross0Degrees || previousAngle > boundaryAngles[1]))
                    {
                        boundaryAngles[0] = previousAngle;
                    }

                    isPreviousClockwise = isCurrentClockwise;
                }
            }

            if (pointAngles[0] == pointAngles[pointAngles.Count - 1] && !hasTurningPoint)
            {
                boundaryAngles[0] = 0;
                boundaryAngles[1] = 360;
            }

            return boundaryAngles;
        }

        private static Drawing.PointF GetVisualCenter(Shape shape)
        {
            Shape duplicateShape = shape.Duplicate()[1];
            duplicateShape.Left = shape.Left;
            duplicateShape.Top = shape.Top;

            PPShape duplicatePPShape = new PPShape(duplicateShape);
            Drawing.PointF visualCenter = duplicatePPShape.VisualCenter;
            duplicateShape.SafeDelete();

            return visualCenter;
        }

        private static bool ListIsPreviouslySelected(List<PPShape> selectedShapes, Dictionary<string, Drawing.PointF> prevSelectedShapes)
        {
            try
            {
                if (selectedShapes == null || selectedShapes.Count <= 0)
                {
                    return false;
                }

                for (int i = 0; i < selectedShapes.Count; i++)
                {
                    Drawing.PointF shapePos = selectedShapes[i].VisualCenter;
                    Drawing.PointF prevShapePos = new Drawing.PointF();
                    if (!prevSelectedShapes.TryGetValue(selectedShapes[i].Name, out prevShapePos))
                    {
                        return false;
                    }

                    if (!(NearlyEqual(shapePos.X, prevShapePos.X, Epsilon) && (NearlyEqual(shapePos.Y, prevShapePos.Y, Epsilon))))
                    {
                        return false;
                    }
                }
            }
            catch
            {
                return false;
            }

            return true;
        }

        private static void SaveSelectedList(List<PPShape> selectedShapes, Dictionary<string, Drawing.PointF> prevSelectedShapes)
        {
            if (selectedShapes == null || selectedShapes.Count <= 0)
            {
                return;
            }

            prevSelectedShapes.Clear();
            for (int i = 0; i < selectedShapes.Count; i++)
            {
                Drawing.PointF shapePos = selectedShapes[i].VisualCenter;
                prevSelectedShapes.Add(selectedShapes[i].Name, shapePos);
            }
        }

        private static List<PPShape> SortPPShapesByName(List<PPShape> selectedShapes, List<string> shapeNames)
        {
            List<PPShape> sortedShapes = new List<PPShape>();
            for (int i = 0; i < shapeNames.Count; i++)
            {
                string name = shapeNames[i];
                for (int j = 0; j < selectedShapes.Count; j++)
                {
                    PPShape shape = selectedShapes[j];
                    if (shape.Name.Equals(name))
                    {
                        sortedShapes.Add(shape);
                        break;
                    }
                }
            }

            return sortedShapes;
        }

        private static List<PPShape> ConvertShapeRangeToPPShapeList(ShapeRange toAlign)
        {
            List<PPShape> selectedShapes = new List<PPShape>();
            for (int i = 1; i <= toAlign.Count; i++)
            {
                Shape s = toAlign[i];
                if (s.Type.Equals(Office.MsoShapeType.msoPicture))
                {
                    selectedShapes.Add(new PPShape(toAlign[i], false));
                }
                else
                {
                    selectedShapes.Add(new PPShape(toAlign[i]));
                }
            }
            return selectedShapes;
        }

        private static void InitDefaultShapesAngles()
        {
            shapeDefaultUpAngle = new Dictionary<MsoAutoShapeType, float>();

            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftArrow, RotateLeft);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftRightArrow, RotateLeft);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftArrowCallout, RotateLeft);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftRightArrowCallout, RotateLeft);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCurvedLeftArrow, RotateLeft);

            shapeDefaultUpAngle.Add(AutoShape.msoShapeRightArrow, RotateRight);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeBentArrow, RotateRight);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeStripedRightArrow, RotateRight);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeNotchedRightArrow, RotateRight);
            shapeDefaultUpAngle.Add(AutoShape.msoShapePentagon, RotateRight);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeChevron, RotateRight);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeRightArrowCallout, RotateRight);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCurvedRightArrow, RotateRight);

            shapeDefaultUpAngle.Add(AutoShape.msoShapeUpArrow, RotateUp);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeBentUpArrow, RotateUp);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeUpDownArrow, RotateUp);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftRightUpArrow, RotateUp);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeLeftUpArrow, RotateUp);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeUpArrowCallout, RotateUp);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCurvedUpArrow, RotateUp);

            shapeDefaultUpAngle.Add(AutoShape.msoShapeDownArrow, RotateDown);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeUTurnArrow, RotateDown);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeDownArrowCallout, RotateDown);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCurvedDownArrow, RotateDown);
            shapeDefaultUpAngle.Add(AutoShape.msoShapeCircularArrow, RotateDown);
        }

        private static Drawing.PointF GetSwapReferencePoint(PPShape shape, PositionsLabSettings.SwapReference r)
        {
            switch (r)
            {
                case PositionsLabSettings.SwapReference.TopLeft:
                    return shape.VisualTopLeft;
                case PositionsLabSettings.SwapReference.TopCenter:
                    return shape.VisualTopCenter;
                case PositionsLabSettings.SwapReference.TopRight:
                    return shape.VisualTopRight;
                case PositionsLabSettings.SwapReference.MiddleLeft:
                    return shape.VisualMiddleLeft;
                case PositionsLabSettings.SwapReference.MiddleCenter:
                    return shape.VisualCenter;
                case PositionsLabSettings.SwapReference.MiddleRight:
                    return shape.VisualMiddleRight;
                case PositionsLabSettings.SwapReference.BottomLeft:
                    return shape.VisualBottomLeft;
                case PositionsLabSettings.SwapReference.BottomCenter:
                    return shape.VisualBottomCenter;
                case PositionsLabSettings.SwapReference.BottomRight:
                    return shape.VisualBottomRight;
                default:
                    return shape.VisualCenter;
            }
        }
        
        private static void InitDefaultAdjoinSettings()
        {
            AdjoinWithAligning();
        }

        private static void InitDefaultSwapSettings()
        {
            prevSelectedShapes = new Dictionary<string, Drawing.PointF>();
            prevSortedShapeNames = new List<string>();
        }

        public static void InitPositionsLab()
        {
            InitDefaultAdjoinSettings();
            InitDefaultSwapSettings();
            InitDefaultShapesAngles();
        }

        #endregion
    }
}
