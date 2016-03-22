using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using AutoShape = Microsoft.Office.Core.MsoAutoShapeType;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using System.Diagnostics;
using Drawing = System.Drawing;

namespace PowerPointLabs.PositionsLab
{
    public class PositionsLabMain
    {
        private const float Epsilon = 0.00001f;
        private const float RotateLeft = 90f;
        private const float RotateRight = 270f;
        private const float RotateUp = 0f;
        private const float RotateDown = 180f;
        private const int None = -1;
        private const int Right = 0;
        private const int Down = 1;
        private const int Left = 2;
        private const int Up = 3;
        private const int Leftorright = 4;
        private const int Upordown = 5;

        //Error Messages
        private const string ErrorMessageFewerThanTwoSelection = TextCollection.PositionsLabText.ErrorFewerThanTwoSelection;
        private const string ErrorMessageFewerThanThreeSelection = TextCollection.PositionsLabText.ErrorFewerThanThreeSelection;
        private const string ErrorMessageUndefined = TextCollection.PositionsLabText.ErrorUndefined;

        public enum DistributeReferenceObject
        {
            Slide,
            FirstShape,
            FirstTwoShapes,
            ExtremeShapes
        }

        public enum DistributeSpaceReferenceObject
        {
            ObjectBoundary,
            ObjectCenter
        }

        //Distribute Grid Variables
        public enum GridAlignment
        {
            AlignLeft,
            AlignCenter,
            AlignRight
        }

        public static DistributeReferenceObject DistributeReference { get; private set; }
        public static DistributeSpaceReferenceObject DistributeSpaceReference { get; private set; }
        public static GridAlignment DistributeGridAlignment { get; private set; }
        public static float MarginTop { get; private set; }
        public static float MarginBottom { get; private set; }
        public static float MarginLeft { get; private set; }
        public static float MarginRight { get; private set; }

        //Reorder Variables
        public enum SwapReference
        {
            TopLeft,
            TopCenter,
            TopRight,
            MiddleLeft,
            MiddleCenter,
            MiddleRight,
            BottomLeft,
            BottomCenter,
            BottomRight
        }

        public static SwapReference SwapReferencePoint { get; set; }

        public static bool IsSwapByClickOrder { get; set; }

        private static Dictionary<int, Drawing.PointF> prevSelectedShapes = new Dictionary<int, Drawing.PointF>();
        private static List<PPShape> prevSortedShapes;

        //Align Variables
        public enum AlignReferenceObject
        {
            Slide,
            SelectedShape,
            PowerpointDefaults
        }
        public static AlignReferenceObject AlignReference { get; private set; }

        // Adjoin Variables
        public static bool AlignShapesToBeAdjoined { get; private set; }

        private static Dictionary<MsoAutoShapeType, float> shapeDefaultUpAngle;

        #region API

        #region Class Methods

        /// <summary>
        /// Tells the Positions Lab to use the slide as the reference point for Align methods
        /// </summary>
        public static void AlignReferToSlide()
        {
            AlignReference = AlignReferenceObject.Slide;
        }

        /// <summary>
        /// Tells the Positions Lab to use first selected shape as reference shape for Align methods
        /// </summary>
        public static void AlignReferToShape()
        {
            AlignReference = AlignReferenceObject.SelectedShape;
        }

        /// <summary>
        /// Tells the Positions Lab to have its align function behave as how powerpoint's does
        /// </summary>
        public static void AlignReferToPowerpointDefaults()
        {
            AlignReference = AlignReferenceObject.PowerpointDefaults;
        }

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

        /// <summary>
        /// Tells the Position Lab to use the slide as the reference point for Distribute methods
        /// </summary>
        public static void DistributeReferToSlide()
        {
            DistributeReference = DistributeReferenceObject.Slide;
        }

        /// <summary>
        /// Tells the Positions Lab to use first selected shape as reference shape for Distribute methods
        /// </summary>
        public static void DistributeReferToFirstShape()
        {
            DistributeReference = DistributeReferenceObject.FirstShape;
        }

        /// <summary>
        /// Tells the Positions Lab to use the first two selected shapes as reference points for Distribute methods
        /// </summary>
        public static void DistributeReferToFirstTwoShapes()
        {
            DistributeReference = DistributeReferenceObject.FirstTwoShapes;
        }

        /// <summary>
        /// Tells the Positions Lab to use detect the corner most shapes and use those as reference points for Distribute methods
        /// </summary>
        public static void DistributeReferToExtremeShapes()
        {
            DistributeReference = DistributeReferenceObject.ExtremeShapes;
        }

        /// <summary>
        /// Tells the Positions Lab to use object boundaries to calculate how much space to distribute the objects by
        /// </summary>
        public static void DistributeSpaceByBoundaries()
        {
            DistributeSpaceReference = DistributeSpaceReferenceObject.ObjectBoundary;
        }

        /// <summary>
        /// Tells the Positions Lab to use the object center to calculate how much space to distribute the objects by
        /// </summary>
        public static void DistributeSpaceByCenter()
        {
            DistributeSpaceReference = DistributeSpaceReferenceObject.ObjectCenter;
        }

        public static void SetDistributeGridAlignment(GridAlignment alignment)
        {
            DistributeGridAlignment = alignment;
        }

        public static void SetDistributeMarginTop (float marginTop)
        {
            MarginTop = marginTop;
        }

        public static void SetDistributeMarginBottom(float marginBottom)
        {
            MarginBottom = marginBottom;
        }

        public static void SetDistributeMarginLeft(float marginLeft)
        {
            MarginLeft = marginLeft;
        }

        public static void SetDistributeMarginRight(float marginRight)
        {
            MarginRight = marginRight;
        }

        #endregion

        #region Align
        public static void AlignLeft(ShapeRange toAlign)
        {
            var selectedShapes = new List<PPShape>();
            for (var i = 1; i <= toAlign.Count; i++)
            {
                selectedShapes.Add(new PPShape(toAlign[i]));
            }

            switch (AlignReference)
            {
                case AlignReferenceObject.Slide: 
                    foreach (var s in selectedShapes)
                    {
                        s.IncrementLeft(-s.Left);
                    }
                    break;
                case AlignReferenceObject.SelectedShape:
                    if (selectedShapes.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    var refShape = selectedShapes[0];

                    for (var i = 1; i < selectedShapes.Count; i++)
                    {
                        var s = selectedShapes[i];
                        s.IncrementLeft(refShape.Left - s.Left);
                    }
                    break;
                case AlignReferenceObject.PowerpointDefaults:
                    if (selectedShapes.Count == 1)
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

            var selectedShapes = new List<PPShape>();
            for (var i = 1; i <= toAlign.Count; i++)
            {
                selectedShapes.Add(new PPShape(toAlign[i]));
            }

            switch (AlignReference)
            {
                case AlignReferenceObject.Slide:
                    foreach (var s in selectedShapes)
                    {
                        s.IncrementLeft(slideWidth - s.Left - s.AbsoluteWidth);
                    }
                    break;
                case AlignReferenceObject.SelectedShape:
                    if (selectedShapes.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    var refShape = selectedShapes[0];
                    var rightMostRefPoint = refShape.Left + refShape.AbsoluteWidth;

                    for (var i = 1; i < selectedShapes.Count; i++)
                    {
                        var s = selectedShapes[i];
                        var rightMostPoint = s.Left + s.AbsoluteWidth;
                        s.IncrementLeft(rightMostRefPoint - rightMostPoint);
                    }
                    break;
                case AlignReferenceObject.PowerpointDefaults:
                    if (selectedShapes.Count == 1)
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
            var selectedShapes = new List<PPShape>();
            for (var i = 1; i <= toAlign.Count; i++)
            {
                selectedShapes.Add(new PPShape(toAlign[i]));
            }

            switch (AlignReference)
            {
                case AlignReferenceObject.Slide:
                    foreach (var s in selectedShapes)
                    {
                        s.IncrementTop(-s.Top);
                    }
                    break;
                case AlignReferenceObject.SelectedShape:
                    if (selectedShapes.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    var refShape = selectedShapes[0];

                    for (var i = 1; i < selectedShapes.Count; i++)
                    {
                        var s = selectedShapes[i];
                        s.IncrementTop(refShape.Top - s.Top);
                    }
                    break;
                case AlignReferenceObject.PowerpointDefaults:
                    if (selectedShapes.Count == 1)
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
            var selectedShapes = new List<PPShape>();
            for (var i = 1; i <= toAlign.Count; i++)
            {
                selectedShapes.Add(new PPShape(toAlign[i]));
            }

            switch (AlignReference)
            {
                case AlignReferenceObject.Slide:
                    foreach (var s in selectedShapes)
                    {
                        s.IncrementTop(slideHeight - s.Top - s.AbsoluteHeight);
                    }
                    break;
                case AlignReferenceObject.SelectedShape:
                    if (selectedShapes.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    var refShape = selectedShapes[0];
                    var lowestRefPoint = refShape.Top + refShape.AbsoluteHeight;

                    for (var i = 1; i < selectedShapes.Count; i++)
                    {
                        var s = selectedShapes[i];
                        var lowestPoint = s.Top + s.AbsoluteHeight;
                        s.IncrementTop(lowestRefPoint - lowestPoint);
                    }
                    break;
                case AlignReferenceObject.PowerpointDefaults:
                    if (selectedShapes.Count == 1)
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
            var selectedShapes = new List<PPShape>();
            for (var i = 1; i <= toAlign.Count; i++)
            {
                selectedShapes.Add(new PPShape(toAlign[i]));
            }

            switch (AlignReference)
            {
                case AlignReferenceObject.Slide:
                    foreach (var s in selectedShapes)
                    {
                        s.IncrementTop(slideHeight / 2 - s.Top - s.AbsoluteHeight / 2);
                    }
                    break;
                case AlignReferenceObject.SelectedShape:
                    if (selectedShapes.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    var refShape = selectedShapes[0];

                    for (var i = 1; i < selectedShapes.Count; i++)
                    {
                        var s = selectedShapes[i];
                        s.IncrementTop(refShape.Center.Y - s.Center.Y);
                    }
                    break;
                case AlignReferenceObject.PowerpointDefaults:
                    if (selectedShapes.Count == 1)
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
            var selectedShapes = new List<PPShape>();
            for (var i = 1; i <= toAlign.Count; i++)
            {
                selectedShapes.Add(new PPShape(toAlign[i]));
            }

            switch (AlignReference)
            {
                case AlignReferenceObject.Slide:
                    foreach (var s in selectedShapes)
                    {
                        s.IncrementLeft(slideWidth / 2 - s.Left - s.AbsoluteWidth / 2);
                    }
                    break;
                case AlignReferenceObject.SelectedShape:
                    if (selectedShapes.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    var refShape = selectedShapes[0];

                    for (var i = 1; i < selectedShapes.Count; i++)
                    {
                        var s = selectedShapes[i];
                        s.IncrementLeft(refShape.Center.X - s.Center.X);
                    }
                    break;
                case AlignReferenceObject.PowerpointDefaults:
                    if (selectedShapes.Count == 1)
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
            var selectedShapes = new List<PPShape>();
            for (var i = 1; i <= toAlign.Count; i++)
            {
                selectedShapes.Add(new PPShape(toAlign[i]));
            }

            switch (AlignReference)
            {
                case AlignReferenceObject.Slide:
                    foreach (var s in selectedShapes)
                    {
                        s.IncrementTop(slideHeight / 2 - s.Top - s.AbsoluteHeight / 2);
                        s.IncrementLeft(slideWidth / 2 - s.Left - s.AbsoluteWidth / 2);
                    }
                    break;
                case AlignReferenceObject.SelectedShape:
                    if (selectedShapes.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    var refShape = selectedShapes[0];

                    for (var i = 1; i < selectedShapes.Count; i++)
                    {
                        var s = selectedShapes[i];
                        s.IncrementTop(refShape.Center.Y - s.Center.Y);
                        s.IncrementLeft(refShape.Center.X - s.Center.X);
                    }
                    break;
                case AlignReferenceObject.PowerpointDefaults:
                    if (selectedShapes.Count == 1)
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

        #endregion

        #region Adjoin
        public static void AdjoinHorizontal(List<PPShape> selectedShapes)
        {
            if (selectedShapes.Count < 2)
            {
                throw new Exception(ErrorMessageFewerThanTwoSelection);
            }

            var refShape = selectedShapes[0];
            var sortedShapes = Graphics.SortShapesByLeft(selectedShapes);
            var refShapeIndex = sortedShapes.IndexOf(refShape);

            var mostLeft = refShape.Left;
            //For all shapes left of refShape, adjoin them from closest to refShape
            for (var i = refShapeIndex - 1; i >= 0; i--)
            {
                var neighbour = sortedShapes[i];
                var rightOfNeighbour = neighbour.Left + neighbour.AbsoluteWidth;
                neighbour.IncrementLeft(mostLeft - rightOfNeighbour);
                if (AlignShapesToBeAdjoined)
                {
                    neighbour.IncrementTop(refShape.Center.Y - neighbour.Center.Y);
                }

                mostLeft = mostLeft - neighbour.AbsoluteWidth;
            }

            var mostRight = refShape.Left + refShape.AbsoluteWidth;
            //For all shapes right of refShape, adjoin them from closest to refShape
            for (var i = refShapeIndex + 1; i < sortedShapes.Count; i++)
            {
                var neighbour = sortedShapes[i];
                neighbour.IncrementLeft(mostRight - neighbour.Left);
                if (AlignShapesToBeAdjoined)
                {
                    neighbour.IncrementTop(refShape.Center.Y - neighbour.Center.Y);
                }

                mostRight = mostRight + neighbour.AbsoluteWidth;
            }
        }

        public static void AdjoinVertical(List<PPShape> selectedShapes)
        {
            if (selectedShapes.Count < 2)
            {
                throw new Exception(ErrorMessageFewerThanTwoSelection);
            }

            var refShape = selectedShapes[0];
            var sortedShapes = Graphics.SortShapesByTop(selectedShapes);
            var refShapeIndex = sortedShapes.IndexOf(refShape);

            var mostTop = refShape.Top;
            //For all shapes above refShape, adjoin them from closest to refShape
            for (var i = refShapeIndex - 1; i >= 0; i--)
            {
                var neighbour = sortedShapes[i];
                var bottomOfNeighbour = neighbour.Top + neighbour.AbsoluteHeight;
                if (AlignShapesToBeAdjoined)
                {
                    neighbour.IncrementLeft(refShape.Center.X - neighbour.Center.X);
                }
                neighbour.IncrementTop(mostTop - bottomOfNeighbour);

                mostTop = mostTop - neighbour.AbsoluteHeight;
            }

            var lowest = refShape.Top + refShape.AbsoluteHeight;
            //For all shapes right of refShape, adjoin them from closest to refShape
            for (var i = refShapeIndex + 1; i < sortedShapes.Count; i++)
            {
                var neighbour = sortedShapes[i];
                if (AlignShapesToBeAdjoined)
                {
                    neighbour.IncrementLeft(refShape.Center.X - neighbour.Center.X);
                }
                neighbour.IncrementTop(lowest - neighbour.Top);

                lowest = lowest + neighbour.AbsoluteHeight;
            }
        }
        #endregion

        #region Distribute
        public static void DistributeHorizontal(List<PPShape> selectedShapes, float slideWidth)
        {
            List<PPShape> sortedShapes;
            var toSortList = new List<PPShape>(selectedShapes);
            PPShape leftRefShape, rightRefShape;
            float startingPoint, width;
            switch (DistributeReference)
            {
                case DistributeReferenceObject.Slide:
                    sortedShapes = Graphics.SortShapesByLeft(selectedShapes);
                    DistributeHorizontal(sortedShapes, slideWidth, 0);
                    break;
                case DistributeReferenceObject.FirstShape:
                    if (selectedShapes.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    var refShape = selectedShapes[0];
                    toSortList.RemoveAt(0);
                    sortedShapes = Graphics.SortShapesByLeft(toSortList);
                    DistributeHorizontal(sortedShapes, refShape.AbsoluteWidth, refShape.Left);
                    break;
                case DistributeReferenceObject.FirstTwoShapes:
                    if (selectedShapes.Count < 3)
                    {
                        throw new Exception(ErrorMessageFewerThanThreeSelection);
                    }

                    leftRefShape = selectedShapes[0];
                    rightRefShape = selectedShapes[1];
                    if (leftRefShape.Left > rightRefShape.Left)
                    {
                        var temp = leftRefShape;
                        leftRefShape = rightRefShape;
                        rightRefShape = temp;
                    }
                    toSortList.RemoveAt(0);
                    toSortList.RemoveAt(0);
                    sortedShapes = Graphics.SortShapesByLeft(toSortList);
                    startingPoint = leftRefShape.Left + leftRefShape.AbsoluteWidth;
                    width = rightRefShape.Left - startingPoint;
                    DistributeHorizontal(sortedShapes, width, startingPoint);
                    break;
                case DistributeReferenceObject.ExtremeShapes:
                    if (selectedShapes.Count < 3)
                    {
                        throw new Exception(ErrorMessageFewerThanThreeSelection);
                    }

                    sortedShapes = Graphics.SortShapesByLeft(toSortList);
                    leftRefShape = sortedShapes[0];
                    rightRefShape = sortedShapes[sortedShapes.Count-1];
                    sortedShapes.RemoveAt(sortedShapes.Count - 1);
                    sortedShapes.RemoveAt(0);
                    startingPoint = leftRefShape.Left + leftRefShape.AbsoluteWidth;
                    width = rightRefShape.Left - startingPoint;
                    DistributeHorizontal(sortedShapes, width, startingPoint);
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private static void DistributeHorizontal(List<PPShape> sortedShapes, float referenceWidth, float startingPoint)
        {
            PPShape refShape;
            var shapeCount = sortedShapes.Count;
            float rightMostRef;
            var spaceBetweenShapes = GetHoritonzalSpaceBetweenShapes(sortedShapes, referenceWidth);

            switch (DistributeSpaceReference)
            {
                case DistributeSpaceReferenceObject.ObjectBoundary:
                    for (var i = 0; i < shapeCount; i++)
                    {
                        var currShape = sortedShapes[i];
                        if (i == 0)
                        {
                            // Check if should have leading space
                            if (spaceBetweenShapes < 0)
                            {
                                currShape.IncrementLeft(startingPoint - currShape.Left);
                            }
                            else
                            {
                                currShape.IncrementLeft(startingPoint - currShape.Left + spaceBetweenShapes);
                            }
                        }
                        else
                        {
                            refShape = sortedShapes[i - 1];
                            rightMostRef = refShape.Left + refShape.AbsoluteWidth;
                            currShape.IncrementLeft(rightMostRef - currShape.Left + spaceBetweenShapes);
                        }
                    }
                    break;
                case DistributeSpaceReferenceObject.ObjectCenter:
                    for (var i =0; i < shapeCount; i++)
                    {
                        var currShape = sortedShapes[i];
                        if (i==0)
                        {
                            currShape.IncrementLeft(startingPoint - currShape.Center.X + spaceBetweenShapes);
                        }
                        else
                        {
                            refShape = sortedShapes[i - 1];
                            rightMostRef = refShape.Center.X;
                            currShape.IncrementLeft(rightMostRef - currShape.Center.X + spaceBetweenShapes);
                        }
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static void DistributeVertical(List<PPShape> selectedShapes, float slideHeight)
        {
            List<PPShape> sortedShapes;
            var toSortList = new List<PPShape>(selectedShapes);
            PPShape topRefShape, bottomRefShape;
            float startingPoint, height;
            switch (DistributeReference)
            {
                case DistributeReferenceObject.Slide:
                    sortedShapes = Graphics.SortShapesByTop(selectedShapes);
                    DistributeVertical(sortedShapes, slideHeight, 0);
                    break;
                case DistributeReferenceObject.FirstShape:
                    if (selectedShapes.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    var refShape = selectedShapes[0];
                    toSortList.RemoveAt(0);
                    sortedShapes = Graphics.SortShapesByTop(toSortList);
                    DistributeVertical(sortedShapes, refShape.AbsoluteHeight, refShape.Top);
                    break;
                case DistributeReferenceObject.FirstTwoShapes:
                    if (selectedShapes.Count < 3)
                    {
                        throw new Exception(ErrorMessageFewerThanThreeSelection);
                    }

                    topRefShape = selectedShapes[0];
                    bottomRefShape = selectedShapes[1];
                    if (topRefShape.Top > bottomRefShape.Top)
                    {
                        var temp = topRefShape;
                        topRefShape = bottomRefShape;
                        bottomRefShape = temp;
                    }
                    toSortList.RemoveAt(0);
                    toSortList.RemoveAt(0);
                    sortedShapes = Graphics.SortShapesByTop(toSortList);
                    startingPoint = topRefShape.Top + topRefShape.AbsoluteHeight;
                    height = bottomRefShape.Top - startingPoint;
                    DistributeVertical(sortedShapes, height, startingPoint);
                    break;
                case DistributeReferenceObject.ExtremeShapes:
                    if (selectedShapes.Count < 3)
                    {
                        throw new Exception(ErrorMessageFewerThanThreeSelection);
                    }

                    sortedShapes = Graphics.SortShapesByTop(selectedShapes);
                    topRefShape = sortedShapes[0];
                    bottomRefShape = sortedShapes[sortedShapes.Count-1];
                    sortedShapes.RemoveAt(sortedShapes.Count - 1);
                    sortedShapes.RemoveAt(0);
                    startingPoint = topRefShape.Top + topRefShape.AbsoluteHeight;
                    height = bottomRefShape.Top - startingPoint;
                    DistributeVertical(sortedShapes, height, startingPoint);
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private static void DistributeVertical(List<PPShape> sortedShapes, float referenceHeight, float startingPoint)
        {
            PPShape refShape;
            var shapeCount = sortedShapes.Count;
            float bottomMostRef;
            var spaceBetweenShapes = GetVerticalSpaceBetweenShapes(sortedShapes, referenceHeight);
      
            switch (DistributeSpaceReference)
            {
                case DistributeSpaceReferenceObject.ObjectBoundary:
                    for (var i = 0; i < shapeCount; i++)
                    {
                        var currShape = sortedShapes[i];
                        if (i == 0)
                        {
                            // Check if should have leading space
                            if (spaceBetweenShapes < 0)
                            {
                                currShape.IncrementTop(startingPoint - currShape.Top);
                            }
                            else
                            {
                                currShape.IncrementTop(startingPoint - currShape.Top + spaceBetweenShapes);
                            }
                        }
                        else
                        {
                            refShape = sortedShapes[i - 1];
                            bottomMostRef = refShape.Top + refShape.AbsoluteHeight;
                            currShape.IncrementTop(bottomMostRef - currShape.Top + spaceBetweenShapes);
                        }
                    }
                    break;
                case DistributeSpaceReferenceObject.ObjectCenter:
                    for (var i = 0; i < shapeCount; i++)
                    {
                        var currShape = sortedShapes[i];
                        if (i == 0)
                        {
                            currShape.IncrementTop(startingPoint - currShape.Center.Y + spaceBetweenShapes);
                        }
                        else
                        {
                            refShape = sortedShapes[i - 1];
                            bottomMostRef = refShape.Center.Y;
                            currShape.IncrementTop(bottomMostRef - currShape.Center.Y + spaceBetweenShapes);
                        }
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static void DistributeCenter(List<PPShape> selectedShapes, float slideWidth, float slideHeight)
        {
            DistributeHorizontal(selectedShapes, slideWidth);
            DistributeVertical(selectedShapes, slideHeight);
        }

        public static void DistributeGrid(List<PPShape> selectedShapes, int rowLength, int colLength)
        {
            var colLengthGivenFullRows = (int) Math.Ceiling((double) selectedShapes.Count/rowLength);
            if (colLength <= colLengthGivenFullRows)
            {
                if (DistributeReference == DistributeReferenceObject.FirstTwoShapes)
                {
                    DistributeGridByRowWithAnchors(selectedShapes, rowLength, colLength);
                }
                else
                {
                    DistributeGridByRow(selectedShapes, rowLength, colLength);
                }
            }
            else
            {
                if (DistributeReference == DistributeReferenceObject.FirstTwoShapes)
                {
                    DistributeGridByColWithAnchors(selectedShapes, rowLength, colLength);
                }
                else
                {
                    DistributeGridByCol(selectedShapes, rowLength, colLength);
                }
            }
        }

        public static void DistributeGridByRow(List<PPShape> selectedShapes, int rowLength, int colLength)
        {
            var refPoint = selectedShapes[0].Center;

            var numShapes = selectedShapes.Count;

            var numIndicesToSkip = IndicesToSkip(numShapes, rowLength, DistributeGridAlignment);

            var rowDifferences = GetLongestWidthsOfRowsByRow(selectedShapes, rowLength, numIndicesToSkip);
            var colDifferences = GetLongestHeightsOfColsByRow(selectedShapes, rowLength, colLength);

            var posX = refPoint.X;
            var posY = refPoint.Y;
            var remainder = numShapes%rowLength;
            var differenceIndex = 0;

            for (var i = 0; i < numShapes; i++)
            {
                //Start of new row
                if (i%rowLength == 0 && i != 0)
                {
                    posX = refPoint.X;
                    differenceIndex = 0;
                    posY += GetSpaceBetweenShapes(i/rowLength - 1, i/rowLength, colDifferences, MarginTop, MarginBottom);
                }

                //If last row, offset by num of indices to skip
                if (numShapes - i == remainder)
                {
                    differenceIndex = numIndicesToSkip;
                    posX += GetSpaceBetweenShapes(0, differenceIndex, rowDifferences, MarginLeft, MarginRight);
                }

                var currentShape = selectedShapes[i];
                currentShape.IncrementLeft(posX - currentShape.Center.X);
                currentShape.IncrementTop(posY - currentShape.Center.Y);

                posX += GetSpaceBetweenShapes(differenceIndex, differenceIndex + 1, rowDifferences, MarginLeft, MarginRight);
                differenceIndex++;
            }
        }

        public static void DistributeGridByCol(List<PPShape> selectedShapes, int rowLength, int colLength)
        {
            var refPoint = selectedShapes[0].Center;

            var numShapes = selectedShapes.Count;

            var numIndicesToSkip = IndicesToSkip(numShapes, colLength, DistributeGridAlignment);

            var rowDifferences = GetLongestWidthsOfRowsByCol(selectedShapes, rowLength, colLength, numIndicesToSkip);
            var colDifferences = GetLongestHeightsOfColsByCol(selectedShapes, rowLength, colLength, numIndicesToSkip);

            var posX = refPoint.X;
            var posY = refPoint.Y;
            var remainder = colLength - (rowLength*colLength - numShapes);
            var augmentedShapeIndex = 0;

            for (var i = 0; i < numShapes; i++)
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
                    posY += GetSpaceBetweenShapes(augmentedShapeIndex/rowLength - 1, augmentedShapeIndex/rowLength, colDifferences, MarginTop, MarginBottom);
                }

                var currentShape = selectedShapes[i];
                var center = currentShape.Center;
                currentShape.IncrementLeft(posX - center.X);
                currentShape.IncrementTop(posY - center.Y);

                posX += GetSpaceBetweenShapes(augmentedShapeIndex%rowLength, augmentedShapeIndex%rowLength + 1, rowDifferences, MarginLeft, MarginRight);
                augmentedShapeIndex++;
            }
        }

        public static void DistributeGridByRowWithAnchors(List<PPShape> selectedShapes, int rowLength, int colLength)
        {
            if (selectedShapes.Count < 2)
            {
                throw new Exception(ErrorMessageFewerThanTwoSelection);
            }

            var startingAnchor = selectedShapes[0].Center;
            var endingAnchor = selectedShapes[1].Center;

            var rowDifference = (endingAnchor.X - startingAnchor.X)/(rowLength - 1);
            var colDifference = (endingAnchor.Y - startingAnchor.Y)/(colLength - 1);

            var numShapes = selectedShapes.Count;
            var numIndicesToSkip = IndicesToSkip(numShapes, rowLength, DistributeGridAlignment);

            var posX = startingAnchor.X;
            var posY = startingAnchor.Y;
            var remainder = numShapes%rowLength;

            var endAnchor = selectedShapes[1];
            selectedShapes.RemoveAt(1);
            selectedShapes.Add(endAnchor);

            for (var i = 0; i < numShapes - 1; i++)
            {
                //Start of new row
                if (i%rowLength == 0 && i != 0)
                {
                    posX = startingAnchor.X;
                    posY += colDifference;
                }

                //If last row, offset by num of indices to skip
                if (numShapes - i == remainder)
                {
                    posX += numIndicesToSkip*rowDifference;
                }

                var currentShape = selectedShapes[i];
                currentShape.IncrementLeft(posX - currentShape.Center.X);
                currentShape.IncrementTop(posY - currentShape.Center.Y);

                posX += rowDifference;
            }
        }

        public static void DistributeGridByColWithAnchors(List<PPShape> selectedShapes, int rowLength, int colLength)
        {
            if (selectedShapes.Count < 2)
            {
                throw new Exception(ErrorMessageFewerThanTwoSelection);
            }

            var startingAnchor = selectedShapes[0].Center;
            var endingAnchor = selectedShapes[1].Center;

            var rowDifference = (endingAnchor.X - startingAnchor.X)/(rowLength - 1);
            var colDifference = (endingAnchor.Y - startingAnchor.Y)/(colLength - 1);

            var numShapes = selectedShapes.Count;

            var numIndicesToSkip = IndicesToSkip(numShapes, colLength, DistributeGridAlignment);

            var posX = startingAnchor.X;
            var posY = startingAnchor.Y;
            var remainder = colLength - (rowLength*colLength - numShapes);
            var augmentedShapeIndex = 0;

            var endAnchor = selectedShapes[1];
            selectedShapes.RemoveAt(1);
            selectedShapes.Add(endAnchor);

            for (var i = 0; i < numShapes - 1; i++)
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
                    posY += colDifference;
                }

                var currentShape = selectedShapes[i];
                var center = currentShape.Center;
                currentShape.IncrementLeft(posX - center.X);
                currentShape.IncrementTop(posY - center.Y);

                posX += rowDifference;
                augmentedShapeIndex++;
            }
        }

        private static float GetHoritonzalSpaceBetweenShapes(List<PPShape> sortedShapes, float referenceWidth)
        {
            var spaceBetweenShapes = referenceWidth;

            switch (DistributeSpaceReference)
            {
                case DistributeSpaceReferenceObject.ObjectBoundary:
                    foreach (var s in sortedShapes)
                    {
                        spaceBetweenShapes -= s.AbsoluteWidth;
                    }
                    break;
                case DistributeSpaceReferenceObject.ObjectCenter:
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            if (spaceBetweenShapes < 0)
            {
                spaceBetweenShapes /= (sortedShapes.Count - 1);
            }
            else
            {
                spaceBetweenShapes /= (sortedShapes.Count + 1);
            }

            return spaceBetweenShapes;
        }

        private static float GetVerticalSpaceBetweenShapes(List<PPShape> sortedShapes, float referenceHeight)
        {
            var spaceBetweenShapes = referenceHeight;

            switch (DistributeSpaceReference)
            {
                case DistributeSpaceReferenceObject.ObjectBoundary:
                    foreach (var s in sortedShapes)
                    {
                        spaceBetweenShapes -= s.AbsoluteHeight;
                    }
                    break;
                case DistributeSpaceReferenceObject.ObjectCenter:
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }

            if (spaceBetweenShapes < 0)
            {
                spaceBetweenShapes /= (sortedShapes.Count - 1);
            }
            else
            {
                spaceBetweenShapes /= (sortedShapes.Count + 1);
            }

            return spaceBetweenShapes;
        }

        #endregion

        #region Swap

        public static void Swap(List<PPShape> selectedShapes)
        {
            if (selectedShapes.Count < 2)
            {
                throw new Exception(ErrorMessageFewerThanTwoSelection);
            }

            var sortedShapes = selectedShapes;

            if (!IsSwapByClickOrder)
            {
                if (ListIsPreviouslySelected(selectedShapes, prevSelectedShapes))
                {
                    sortedShapes = prevSortedShapes;
                }
                else
                {
                    sortedShapes = Graphics.SortShapesByLeft(selectedShapes);
                }
            }
            else
            {
                prevSelectedShapes.Clear();
            }

            var firstPos = GetSwapReferencePoint(sortedShapes[0], SwapReferencePoint);

            prevSortedShapes = new List<PPShape>();

            for (var i = 0; i < sortedShapes.Count; i++)
            {
                var currentShape = sortedShapes[i];
                if (i < sortedShapes.Count - 1)
                {
                    var currentPos = GetSwapReferencePoint(currentShape, SwapReferencePoint);
                    var nextPos = GetSwapReferencePoint(sortedShapes[i + 1], SwapReferencePoint);

                    currentShape.IncrementLeft(nextPos.X - currentPos.X);
                    currentShape.IncrementTop(nextPos.Y - currentPos.Y);
                }
                else
                {
                    var currentPos = GetSwapReferencePoint(currentShape, SwapReferencePoint);
                    currentShape.IncrementLeft(firstPos.X - currentPos.X);
                    currentShape.IncrementTop(firstPos.Y - currentPos.Y);
                }

                if (i != 0 && !IsSwapByClickOrder)
                {
                    prevSortedShapes.Add(currentShape);
                }
            }

            if (!IsSwapByClickOrder)
            {
                prevSortedShapes.Add(sortedShapes[0]);
                SaveSelectedList(prevSortedShapes, prevSelectedShapes);
            }
        }

        #endregion

        #region Snap

        public static void SnapVertical(List<Shape> selectedShapes)
        {
            foreach (var s in selectedShapes)
            {
                SnapShapeVertical(s);
            }
        }

        public static void SnapHorizontal(List<Shape> selectedShapes)
        {
            foreach (var s in selectedShapes)
            {
                SnapShapeHorizontal(s);
            }
        }

        public static void SnapAway(List<Shape> shapes)
        {
            if (shapes.Count < 2)
            {
                throw new Exception(ErrorMessageFewerThanTwoSelection);
            }

            var refShapeCenter = Graphics.GetCenterPoint(shapes[0]);
            var isAllSameDir = true;
            var lastDir = -1;

            for (var i = 1; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var shapeCenter = Graphics.GetCenterPoint(shape);
                var angle = (float) AngleBetweenTwoPoints(refShapeCenter, shapeCenter);

                var dir = GetDirectionWrtRefShape(shape, angle);

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
                if (dir < Leftorright)
                {
                    lastDir = dir;
                }
            }

            if (!isAllSameDir || lastDir == None)
            {
                lastDir = 0;
            }
            else
            {
                lastDir++;
            }

            for (var i = 1; i < shapes.Count; i++)
            {
                var shape = shapes[i];
                var shapeCenter = Graphics.GetCenterPoint(shape);
                var angle = (float) AngleBetweenTwoPoints(refShapeCenter, shapeCenter);

                float defaultUpAngle = 0;
                var hasDefaultDirection = shapeDefaultUpAngle.TryGetValue(shape.AutoShapeType, out defaultUpAngle);

                if (hasDefaultDirection)
                {
                    shape.Rotation = (defaultUpAngle + angle) + lastDir*90;
                }
                else
                {
                    if (IsVertical(shape))
                    {
                        shape.Rotation = angle + lastDir*90;
                    }
                    else
                    {
                        shape.Rotation = (angle - 90) + lastDir*90;
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

        private static void SnapTo0Or180(Shape shape)
        {
            var rotation = shape.Rotation;

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
            var rotation = shape.Rotation;

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

        #endregion

        #region Util

        public static double AngleBetweenTwoPoints(Drawing.PointF refPoint, Drawing.PointF pt)
        {
            var angle = Math.Atan((pt.Y - refPoint.Y)/(pt.X - refPoint.X))*180/Math.PI;

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
            var absA = Math.Abs(a);
            var absB = Math.Abs(b);
            var diff = Math.Abs(a - b);

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
            var hasDefaultDirection = shapeDefaultUpAngle.TryGetValue(shape.AutoShapeType, out defaultUpAngle);

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

            var angle = AddAngles(angleFromRefShape, defaultUpAngle);
            var diff = SubtractAngles(shape.Rotation, angle);
            var phaseInFloat = diff/90;

            if (shape.AutoShapeType == AutoShape.msoShapeLightningBolt)
            {
                Debug.WriteLine("angle: " + angle);
                Debug.WriteLine("diff: " + diff);
                Debug.WriteLine("phaseInFloat: " + defaultUpAngle);
                Debug.WriteLine("equal: " + NearlyEqual(phaseInFloat, (float) Math.Round(phaseInFloat), Epsilon));
            }

            if (!NearlyEqual(phaseInFloat, (float) Math.Round(phaseInFloat), Epsilon))
            {
                return None;
            }

            var phase = (int) Math.Round(phaseInFloat);

            if (!hasDefaultDirection)
            {
                if (phase == Left || phase == Right)
                {
                    return Leftorright;
                }

                return Upordown;
            }

            return phase;
        }

        private static bool IsSameDirection(int a, int b)
        {
            if (a == b) return true;
            if (a == Leftorright) return b == Left || b == Right;
            if (b == Leftorright) return a == Left || a == Right;
            if (a == Upordown) return b == Up || b == Down;
            if (b == Upordown) return a == Up || a == Down;

            return false;
        }

        public static float AddAngles(float a, float b)
        {
            return (a + b)%360;
        }

        public static float SubtractAngles(float a, float b)
        {
            var diff = a - b;
            if (diff < 0)
            {
                return 360 + diff;
            }

            return diff;
        }

        public static float[] GetLongestWidthsOfRowsByRow(List<PPShape> shapes, int rowLength, int numIndicesToSkip)
        {
            var longestWidths = new float[rowLength];
            var numShapes = shapes.Count;
            var remainder = numShapes%rowLength;

            for (var i = 0; i < numShapes; i++)
            {
                var longestRowIndex = i%rowLength;
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

        public static float[] GetLongestHeightsOfColsByRow(List<PPShape> shapes, int rowLength, int colLength)
        {
            var longestHeights = new float[colLength];

            for (var i = 0; i < shapes.Count; i++)
            {
                var longestHeightIndex = i/rowLength;
                if (longestHeights[longestHeightIndex] < shapes[i].AbsoluteHeight)
                {
                    longestHeights[longestHeightIndex] = shapes[i].AbsoluteHeight;
                }
            }

            return longestHeights;
        }

        public static float[] GetLongestWidthsOfRowsByCol(List<PPShape> shapes, int rowLength, int colLength, int numIndicesToSkip)
        {
            var longestWidths = new float[rowLength];
            var numShapes = shapes.Count;
            var augmentedShapeIndex = 0;
            var remainder = colLength - (rowLength*colLength - numShapes);

            for (var i = 0; i < numShapes; i++)
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

                var longestWidthsArrayIndex = augmentedShapeIndex%rowLength;

                if (longestWidths[longestWidthsArrayIndex] < shapes[i].AbsoluteWidth)
                {
                    longestWidths[longestWidthsArrayIndex] = shapes[i].AbsoluteWidth;
                }

                augmentedShapeIndex++;
            }

            return longestWidths;
        }

        public static float[] GetLongestHeightsOfColsByCol(List<PPShape> shapes, int rowLength, int colLength, int numIndicesToSkip)
        {
            var longestHeights = new float[colLength];
            var numShapes = shapes.Count;
            var augmentedShapeIndex = 0;
            var remainder = colLength - (rowLength*colLength - numShapes);

            for (var i = 0; i < numShapes; i++)
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

                var longestHeightArrayIndex = augmentedShapeIndex/rowLength;

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

        public static int IndicesToSkip(int totalSelectedShapes, int rowLength, GridAlignment alignment)
        {
            var numOfShapesInLastRow = totalSelectedShapes%rowLength;

            if (alignment == GridAlignment.AlignLeft || numOfShapesInLastRow == 0)
            {
                return 0;
            }

            if (alignment == GridAlignment.AlignRight)
            {
                return rowLength - numOfShapesInLastRow;
            }

            if (alignment == GridAlignment.AlignCenter)
            {
                var difference = rowLength - numOfShapesInLastRow;
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

            var start = 0;
            var end = 0;

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

            for (var i = start; i < end; i++)
            {
                difference += (differences[i] / 2 + margin1 + margin2 + differences[i + 1] / 2);
            }

            return difference;
        }

        private static bool ListIsPreviouslySelected(List<PPShape> selectedShapes, Dictionary<int, Drawing.PointF> prevSelectedShapes)
        {
            try
            {
                if (selectedShapes == null || selectedShapes.Count <= 0)
                {
                    return false;
                }

                for (int i = 0; i < selectedShapes.Count; i++)
                {
                    var shapePos = selectedShapes[i].Center;
                    Drawing.PointF prevShapePos = new Drawing.PointF();
                    if (!prevSelectedShapes.TryGetValue(selectedShapes[i].Id, out prevShapePos))
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

        private static void SaveSelectedList(List<PPShape> selectedShapes, Dictionary<int, Drawing.PointF> prevSelectedShapes)
        {
            if (selectedShapes == null || selectedShapes.Count <= 0)
            {
                return;
            }

            prevSelectedShapes.Clear();
            for (int i = 0; i < selectedShapes.Count; i++)
            {
                var shapePos = selectedShapes[i].Center;
                prevSelectedShapes.Add(selectedShapes[i].Id, shapePos);
            }
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

        private static void InitDefaultDistributeSettings()
        {
            MarginTop = 5;
            MarginBottom = 5;
            MarginLeft = 5;
            MarginRight = 5;
            DistributeGridAlignment = GridAlignment.AlignLeft;
            DistributeReferToFirstTwoShapes();
            DistributeSpaceByBoundaries();
        }

        private static Drawing.PointF GetSwapReferencePoint(PPShape shape, SwapReference r)
        {
            switch (r)
            {
                case SwapReference.TopLeft:
                    return shape.TopLeft;
                case SwapReference.TopCenter:
                    return shape.TopCenter;
                case SwapReference.TopRight:
                    return shape.TopRight;
                case SwapReference.MiddleLeft:
                    return shape.MiddleLeft;
                case SwapReference.MiddleCenter:
                    return shape.Center;
                case SwapReference.MiddleRight:
                    return shape.MiddleRight;
                case SwapReference.BottomLeft:
                    return shape.BottomLeft;
                case SwapReference.BottomCenter:
                    return shape.BottomCenter;
                case SwapReference.BottomRight:
                    return shape.BottomRight;
                default:
                    return shape.Center;
            }
        }

        private static void InitDefaultSwapSettings()
        {
            IsSwapByClickOrder = false;
            SwapReferencePoint = SwapReference.MiddleCenter;
        }

        private static void InitDefaultAlignSettings()
        {
            AlignReferToShape();
        }

        private static void InitDefaultAdjoinSettings()
        {
            AdjoinWithAligning();
        }

        public static void InitPositionsLab()
        {
            InitDefaultShapesAngles();
            InitDefaultDistributeSettings();
            InitDefaultAlignSettings();
            InitDefaultSwapSettings();
            InitDefaultAdjoinSettings();
        }

        #endregion
    }
}
