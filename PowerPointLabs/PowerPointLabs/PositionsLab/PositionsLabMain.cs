using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using AutoShape = Microsoft.Office.Core.MsoAutoShapeType;
using ShapeRange = Microsoft.Office.Interop.PowerPoint.ShapeRange;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using Drawing = System.Drawing;
using System.Linq;

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

        //Error Messages
        private const string ErrorMessageNoSelection = TextCollection.PositionsLabText.ErrorNoSelection;
        private const string ErrorMessageFewerThanTwoSelection = TextCollection.PositionsLabText.ErrorFewerThanTwoSelection;
        private const string ErrorMessageFewerThanThreeSelection = TextCollection.PositionsLabText.ErrorFewerThanThreeSelection;
        private const string ErrorMessageFewerThanFourSelection = TextCollection.PositionsLabText.ErrorFewerThanFourSelection;
        private const string ErrorMessageFunctionNotSupportedForExtremeShapes =
            TextCollection.PositionsLabText.ErrorFunctionNotSupportedForWithinShapes;
        private const string ErrorMessageFunctionNotSupportedForSlide =
            TextCollection.PositionsLabText.ErrorFunctionNotSupportedForSlide;
        private const string ErrorMessageFunctionNotSuppertedForOverlapRefShapeCenter =
            TextCollection.PositionsLabText.ErrorFunctionNotSupportedForOverlapRefShapeCenter;
        private const string ErrorMessageUndefined = TextCollection.PositionsLabText.ErrorUndefined;

        public enum DistributeReferenceObject
        {
            Slide,
            FirstShape,
            FirstTwoShapes,
            ExtremeShapes
        }

        public enum DistributeRadialReferenceObject
        {
            AtSecondShape,
            SecondThirdShape
        }

        public enum DistributeSpaceReferenceObject
        {
            ObjectBoundary,
            ObjectCenter
        }

        //Distribute Grid Variables
        public enum GridAlignment
        {
            None,
            AlignLeft,
            AlignCenter,
            AlignRight
        }

        public static DistributeReferenceObject DistributeReference { get; private set; }
        public static DistributeRadialReferenceObject DistributeRadialReference { get; private set; }
        public static DistributeSpaceReferenceObject DistributeSpaceReference { get; private set; }
        public static GridAlignment DistributeGridAlignment { get; private set; }
        public static float MarginTop { get; private set; }
        public static float MarginBottom { get; private set; }
        public static float MarginLeft { get; private set; }
        public static float MarginRight { get; private set; }

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

        private static Dictionary<string, Drawing.PointF> prevSelectedShapes = new Dictionary<string, Drawing.PointF>();
        private static List<string> prevSortedShapeNames;

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

        // Radial Variables
        public enum RadialShapeOrientationObject
        {
            Fixed,
            Dynamic
        }
        public static RadialShapeOrientationObject DistributeShapeOrientation { get; private set; }
        public static RadialShapeOrientationObject ReorientShapeOrientation { get; private set; }

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
        /// Tells the Position Lab to use the second selected shape as the starting point for Distribute Radial method
        /// </summary>
        public static void DistributeReferAtSecondShape()
        {
            DistributeRadialReference = DistributeRadialReferenceObject.AtSecondShape;
        }

        /// <summary>
        /// Tells the Position Lab to use the second and third shape as the boundary points for Distribute Radial method
        /// </summary>
        public static void DistributeReferToSecondThirdShape()
        {
            DistributeRadialReference = DistributeRadialReferenceObject.SecondThirdShape;
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

        public static void DistributeShapeOrientationToFixed()
        {
            DistributeShapeOrientation = RadialShapeOrientationObject.Fixed;
        }

        public static void DistributeShapeOrientationToDynamic()
        {
            DistributeShapeOrientation = RadialShapeOrientationObject.Dynamic;
        }

        public static void ReorientShapeOrientationToFixed()
        {
            ReorientShapeOrientation = RadialShapeOrientationObject.Fixed;
        }

        public static void ReorientShapeOrientationToDynamic()
        {
            ReorientShapeOrientation = RadialShapeOrientationObject.Dynamic;
        }

        #endregion

        #region Align
        public static void AlignLeft(ShapeRange toAlign)
        {
            var selectedShapes = new List<PPShape>();

            switch (AlignReference)
            {
                case AlignReferenceObject.Slide:
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    foreach (var s in selectedShapes)
                    {
                        s.IncrementLeft(-s.VisualLeft);
                    }
                    break;
                case AlignReferenceObject.SelectedShape:
                    if (toAlign.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    var refShape = selectedShapes[0];

                    for (var i = 1; i < selectedShapes.Count; i++)
                    {
                        var s = selectedShapes[i];
                        s.IncrementLeft(refShape.VisualLeft - s.VisualLeft);
                    }
                    break;
                case AlignReferenceObject.PowerpointDefaults:

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

            var selectedShapes = new List<PPShape>();

            switch (AlignReference)
            {
                case AlignReferenceObject.Slide:
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    foreach (var s in selectedShapes)
                    {
                        s.IncrementLeft(slideWidth - s.VisualLeft - s.AbsoluteWidth);
                    }
                    break;
                case AlignReferenceObject.SelectedShape:
                    if (toAlign.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    var refShape = selectedShapes[0];
                    var rightMostRefPoint = refShape.VisualLeft + refShape.AbsoluteWidth;

                    for (var i = 1; i < selectedShapes.Count; i++)
                    {
                        var s = selectedShapes[i];
                        var rightMostPoint = s.VisualLeft + s.AbsoluteWidth;
                        s.IncrementLeft(rightMostRefPoint - rightMostPoint);
                    }
                    break;
                case AlignReferenceObject.PowerpointDefaults:
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
            var selectedShapes = new List<PPShape>();

            switch (AlignReference)
            {
                case AlignReferenceObject.Slide:
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    foreach (var s in selectedShapes)
                    {
                        s.IncrementTop(-s.VisualTop);
                    }
                    break;
                case AlignReferenceObject.SelectedShape:
                    if (toAlign.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    var refShape = selectedShapes[0];

                    for (var i = 1; i < selectedShapes.Count; i++)
                    {
                        var s = selectedShapes[i];
                        s.IncrementTop(refShape.VisualTop - s.VisualTop);
                    }
                    break;
                case AlignReferenceObject.PowerpointDefaults:
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
            var selectedShapes = new List<PPShape>();

            switch (AlignReference)
            {
                case AlignReferenceObject.Slide:
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    foreach (var s in selectedShapes)
                    {
                        s.IncrementTop(slideHeight - s.VisualTop - s.AbsoluteHeight);
                    }
                    break;
                case AlignReferenceObject.SelectedShape:
                    if (toAlign.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    var refShape = selectedShapes[0];
                    var lowestRefPoint = refShape.VisualTop + refShape.AbsoluteHeight;

                    for (var i = 1; i < selectedShapes.Count; i++)
                    {
                        var s = selectedShapes[i];
                        var lowestPoint = s.VisualTop + s.AbsoluteHeight;
                        s.IncrementTop(lowestRefPoint - lowestPoint);
                    }
                    break;
                case AlignReferenceObject.PowerpointDefaults:
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
            var selectedShapes = new List<PPShape>();

            switch (AlignReference)
            {
                case AlignReferenceObject.Slide:
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    foreach (var s in selectedShapes)
                    {
                        s.IncrementTop(slideHeight / 2 - s.VisualTop - s.AbsoluteHeight / 2);
                    }
                    break;
                case AlignReferenceObject.SelectedShape:
                    if (toAlign.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    var refShape = selectedShapes[0];

                    for (var i = 1; i < selectedShapes.Count; i++)
                    {
                        var s = selectedShapes[i];
                        s.IncrementTop(refShape.VisualCenter.Y - s.VisualCenter.Y);
                    }
                    break;
                case AlignReferenceObject.PowerpointDefaults:
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
            var selectedShapes = new List<PPShape>();

            switch (AlignReference)
            {
                case AlignReferenceObject.Slide:
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    foreach (var s in selectedShapes)
                    {
                        s.IncrementLeft(slideWidth / 2 - s.VisualLeft - s.AbsoluteWidth / 2);
                    }
                    break;
                case AlignReferenceObject.SelectedShape:
                    if (toAlign.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    var refShape = selectedShapes[0];

                    for (var i = 1; i < selectedShapes.Count; i++)
                    {
                        var s = selectedShapes[i];
                        s.IncrementLeft(refShape.VisualCenter.X - s.VisualCenter.X);
                    }
                    break;
                case AlignReferenceObject.PowerpointDefaults:
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
            var selectedShapes = new List<PPShape>();

            switch (AlignReference)
            {
                case AlignReferenceObject.Slide:
                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    foreach (var s in selectedShapes)
                    {
                        s.IncrementTop(slideHeight / 2 - s.VisualTop - s.AbsoluteHeight / 2);
                        s.IncrementLeft(slideWidth / 2 - s.VisualLeft - s.AbsoluteWidth / 2);
                    }
                    break;
                case AlignReferenceObject.SelectedShape:
                    if (toAlign.Count < 2)
                    {
                        throw new Exception(ErrorMessageFewerThanTwoSelection);
                    }

                    selectedShapes = ConvertShapeRangeToPPShapeList(toAlign);
                    var refShape = selectedShapes[0];

                    for (var i = 1; i < selectedShapes.Count; i++)
                    {
                        var s = selectedShapes[i];
                        s.IncrementTop(refShape.VisualCenter.Y - s.VisualCenter.Y);
                        s.IncrementLeft(refShape.VisualCenter.X - s.VisualCenter.X);
                    }
                    break;
                case AlignReferenceObject.PowerpointDefaults:
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
                throw new Exception(ErrorMessageFewerThanThreeSelection);
            }
                    
            var origin = Graphics.GetCenterPoint(selectedShapes[1]);
            var refPoint = Graphics.GetCenterPoint(selectedShapes[2]);
            var distance = DistanceBetweenTwoPoints(origin, refPoint);

            for (var i = 3; i <= selectedShapes.Count; i++)
            {
                var shape = selectedShapes[i];
                var point = Graphics.GetCenterPoint(shape);
                var currentDistance = DistanceBetweenTwoPoints(origin, point);
                var proportion = (currentDistance - distance) / currentDistance;

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
                throw new Exception(ErrorMessageFewerThanTwoSelection);
            }

            var refShape = selectedShapes[0];
            var sortedShapes = Graphics.SortShapesByLeft(selectedShapes);
            var refShapeIndex = sortedShapes.IndexOf(refShape);

            var mostLeft = refShape.VisualLeft;
            //For all shapes left of refShape, adjoin them from closest to refShape
            for (var i = refShapeIndex - 1; i >= 0; i--)
            {
                var neighbour = sortedShapes[i];
                var rightOfNeighbour = neighbour.VisualLeft + neighbour.AbsoluteWidth;
                neighbour.IncrementLeft(mostLeft - rightOfNeighbour);
                if (AlignShapesToBeAdjoined)
                {
                    neighbour.IncrementTop(refShape.VisualCenter.Y - neighbour.VisualCenter.Y);
                }

                mostLeft = mostLeft - neighbour.AbsoluteWidth;
            }

            var mostRight = refShape.VisualLeft + refShape.AbsoluteWidth;
            //For all shapes right of refShape, adjoin them from closest to refShape
            for (var i = refShapeIndex + 1; i < sortedShapes.Count; i++)
            {
                var neighbour = sortedShapes[i];
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
                throw new Exception(ErrorMessageFewerThanTwoSelection);
            }

            var refShape = selectedShapes[0];
            var sortedShapes = Graphics.SortShapesByTop(selectedShapes);
            var refShapeIndex = sortedShapes.IndexOf(refShape);

            var mostTop = refShape.VisualTop;
            //For all shapes above refShape, adjoin them from closest to refShape
            for (var i = refShapeIndex - 1; i >= 0; i--)
            {
                var neighbour = sortedShapes[i];
                var bottomOfNeighbour = neighbour.VisualTop + neighbour.AbsoluteHeight;
                if (AlignShapesToBeAdjoined)
                {
                    neighbour.IncrementLeft(refShape.VisualCenter.X - neighbour.VisualCenter.X);
                }
                neighbour.IncrementTop(mostTop - bottomOfNeighbour);

                mostTop = mostTop - neighbour.AbsoluteHeight;
            }

            var lowest = refShape.VisualTop + refShape.AbsoluteHeight;
            //For all shapes right of refShape, adjoin them from closest to refShape
            for (var i = refShapeIndex + 1; i < sortedShapes.Count; i++)
            {
                var neighbour = sortedShapes[i];
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
            var isSlide = DistributeReference == DistributeReferenceObject.Slide;
            var isFirstShape = DistributeReference == DistributeReferenceObject.FirstShape;
            var isExtremeShape = DistributeReference == DistributeReferenceObject.ExtremeShapes;
            var isFirstTwoShapes = DistributeReference == DistributeReferenceObject.FirstTwoShapes;
            var isObjectCenter = DistributeSpaceReference == DistributeSpaceReferenceObject.ObjectCenter;
            var isObjectBoundary = DistributeSpaceReference == DistributeSpaceReferenceObject.ObjectBoundary;

            List<PPShape> shapesToDistribute;
            PPShape refShape;
            float referenceWidth, spaceBetweenShapes, startingPoint, rightMostRef, totalShapeWidth = 0;

            if (isSlide && isObjectBoundary)
            {
                if (selectedShapes.Count < 1)
                {
                    throw new Exception(ErrorMessageNoSelection);
                }

                startingPoint = 0;
                referenceWidth = slideWidth;
                shapesToDistribute = Graphics.SortShapesByLeft(selectedShapes);

                foreach (var s in shapesToDistribute)
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

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];

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
                    throw new Exception(ErrorMessageNoSelection);
                }

                startingPoint = 0;
                referenceWidth = slideWidth;
                shapesToDistribute = Graphics.SortShapesByLeft(selectedShapes);

                foreach (var s in shapesToDistribute)
                {
                    totalShapeWidth += s.AbsoluteWidth;
                }

                if (totalShapeWidth > referenceWidth)
                {
                    var leftMostShape = shapesToDistribute[0];
                    var rightMostShape = shapesToDistribute[shapesToDistribute.Count - 1];

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

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];
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
                    throw new Exception(ErrorMessageFewerThanTwoSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualLeft;
                referenceWidth = shapesToDistribute[0].AbsoluteWidth;
                shapesToDistribute.RemoveAt(0);
                shapesToDistribute = Graphics.SortShapesByLeft(shapesToDistribute);

                foreach (var s in shapesToDistribute)
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

                for (var i =0; i <shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];

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
                    throw new Exception(ErrorMessageFewerThanTwoSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualLeft;
                referenceWidth = shapesToDistribute[0].AbsoluteWidth;
                shapesToDistribute.RemoveAt(0);
                shapesToDistribute = Graphics.SortShapesByLeft(shapesToDistribute);

                foreach (var s in shapesToDistribute)
                {
                    totalShapeWidth += s.AbsoluteWidth;
                }

                if (totalShapeWidth > referenceWidth)
                {
                    var leftMostShape = shapesToDistribute[0];
                    var rightMostShape = shapesToDistribute[shapesToDistribute.Count - 1];

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

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];
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
                    throw new Exception(ErrorMessageFewerThanThreeSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);

                if (shapesToDistribute[0].VisualLeft > shapesToDistribute[1].VisualLeft)
                {
                    var temp = shapesToDistribute[0];
                    shapesToDistribute[0] = shapesToDistribute[1];
                    shapesToDistribute[1] = temp;
                }

                startingPoint = shapesToDistribute[0].VisualLeft + shapesToDistribute[0].AbsoluteWidth;
                referenceWidth = shapesToDistribute[1].VisualLeft + shapesToDistribute[1].AbsoluteWidth - shapesToDistribute[0].VisualLeft;

                foreach (var s in shapesToDistribute)
                {
                    totalShapeWidth += s.AbsoluteWidth;
                }
                spaceBetweenShapes = (referenceWidth - totalShapeWidth) / (shapesToDistribute.Count - 1);

                shapesToDistribute.RemoveAt(1);
                shapesToDistribute.RemoveAt(0);
                shapesToDistribute = Graphics.SortShapesByLeft(shapesToDistribute);

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];
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
                    throw new Exception(ErrorMessageFewerThanThreeSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);

                if (shapesToDistribute[0].VisualLeft > shapesToDistribute[1].VisualLeft)
                {
                    var temp = shapesToDistribute[0];
                    shapesToDistribute[0] = shapesToDistribute[1];
                    shapesToDistribute[1] = temp;
                }

                startingPoint = shapesToDistribute[0].VisualCenter.X;
                referenceWidth = shapesToDistribute[1].VisualCenter.X -shapesToDistribute[0].VisualCenter.X;

                foreach (var s in shapesToDistribute)
                {
                    totalShapeWidth += s.AbsoluteWidth;
                }
                
                spaceBetweenShapes = referenceWidth / (shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(1);
                shapesToDistribute.RemoveAt(0);

                shapesToDistribute = Graphics.SortShapesByLeft(shapesToDistribute);

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];
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
                    throw new Exception(ErrorMessageFewerThanThreeSelection);
                }

                shapesToDistribute = Graphics.SortShapesByLeft(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualLeft + shapesToDistribute[0].AbsoluteWidth;
                var rightMostShape = shapesToDistribute[shapesToDistribute.Count - 1];
                referenceWidth = rightMostShape.VisualLeft + rightMostShape.AbsoluteWidth - shapesToDistribute[0].VisualLeft;

                foreach (var s in shapesToDistribute)
                {
                    totalShapeWidth += s.AbsoluteWidth;
                }
                spaceBetweenShapes = (referenceWidth - totalShapeWidth) / (shapesToDistribute.Count - 1);

                shapesToDistribute.RemoveAt(shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(0);

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];
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
                    throw new Exception(ErrorMessageFewerThanThreeSelection);
                }

                shapesToDistribute = Graphics.SortShapesByLeft(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualCenter.X;
                referenceWidth = shapesToDistribute[shapesToDistribute.Count-1].VisualCenter.X - shapesToDistribute[0].VisualCenter.X;

                foreach (var s in shapesToDistribute)
                {
                    totalShapeWidth += s.AbsoluteWidth;
                }

                spaceBetweenShapes = referenceWidth / (shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(0);

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];
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
            var isSlide = DistributeReference == DistributeReferenceObject.Slide;
            var isFirstShape = DistributeReference == DistributeReferenceObject.FirstShape;
            var isExtremeShape = DistributeReference == DistributeReferenceObject.ExtremeShapes;
            var isFirstTwoShapes = DistributeReference == DistributeReferenceObject.FirstTwoShapes;
            var isObjectCenter = DistributeSpaceReference == DistributeSpaceReferenceObject.ObjectCenter;
            var isObjectBoundary = DistributeSpaceReference == DistributeSpaceReferenceObject.ObjectBoundary;

            List<PPShape> shapesToDistribute;
            PPShape refShape;
            float referenceHeight, spaceBetweenShapes, startingPoint, bottomMostRef, totalShapeHeight = 0;

            if (isSlide && isObjectBoundary)
            {
                if (selectedShapes.Count < 1)
                {
                    throw new Exception(ErrorMessageNoSelection);
                }

                startingPoint = 0;
                referenceHeight = slideHeight;
                shapesToDistribute = Graphics.SortShapesByTop(selectedShapes);
                foreach (var s in shapesToDistribute)
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

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];

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
                    throw new Exception(ErrorMessageNoSelection);
                }

                startingPoint = 0;
                referenceHeight = slideHeight;
                shapesToDistribute = Graphics.SortShapesByTop(selectedShapes);

                foreach (var s in shapesToDistribute)
                {
                    totalShapeHeight += s.AbsoluteHeight;
                }

                if (totalShapeHeight > referenceHeight)
                {
                    var topMostShape = shapesToDistribute[0];
                    var bottomMostShape = shapesToDistribute[shapesToDistribute.Count - 1];

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

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];
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
                    throw new Exception(ErrorMessageFewerThanTwoSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualTop;
                referenceHeight = shapesToDistribute[0].AbsoluteHeight;
                shapesToDistribute.RemoveAt(0);
                shapesToDistribute = Graphics.SortShapesByTop(shapesToDistribute);

                foreach (var s in shapesToDistribute)
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

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];

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
                    throw new Exception(ErrorMessageFewerThanTwoSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualTop;
                referenceHeight = shapesToDistribute[0].AbsoluteHeight;
                shapesToDistribute.RemoveAt(0);
                shapesToDistribute = Graphics.SortShapesByTop(shapesToDistribute);

                foreach (var s in shapesToDistribute)
                {
                    totalShapeHeight += s.AbsoluteHeight;
                }

                if (totalShapeHeight > referenceHeight)
                {
                    var topMostShape = shapesToDistribute[0];
                    var bottomMostShape = shapesToDistribute[shapesToDistribute.Count - 1];

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

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];
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
                    throw new Exception(ErrorMessageFewerThanThreeSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);

                if (shapesToDistribute[0].VisualTop > shapesToDistribute[1].VisualTop)
                {
                    var temp = shapesToDistribute[0];
                    shapesToDistribute[0] = shapesToDistribute[1];
                    shapesToDistribute[1] = temp;
                }

                startingPoint = shapesToDistribute[0].VisualTop + shapesToDistribute[0].AbsoluteHeight;
                referenceHeight = shapesToDistribute[1].VisualTop + shapesToDistribute[1].AbsoluteHeight - shapesToDistribute[0].VisualTop;

                foreach (var s in shapesToDistribute)
                {
                    totalShapeHeight += s.AbsoluteHeight;
                }
                spaceBetweenShapes = (referenceHeight - totalShapeHeight) / (shapesToDistribute.Count - 1);

                shapesToDistribute.RemoveAt(1);
                shapesToDistribute.RemoveAt(0);
                shapesToDistribute = Graphics.SortShapesByTop(shapesToDistribute);

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];
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
                    throw new Exception(ErrorMessageFewerThanThreeSelection);
                }

                shapesToDistribute = new List<PPShape>(selectedShapes);

                if (shapesToDistribute[0].VisualTop > shapesToDistribute[1].VisualTop)
                {
                    var temp = shapesToDistribute[0];
                    shapesToDistribute[0] = shapesToDistribute[1];
                    shapesToDistribute[1] = temp;
                }

                startingPoint = shapesToDistribute[0].VisualCenter.Y;
                referenceHeight = shapesToDistribute[1].VisualCenter.Y - shapesToDistribute[0].VisualCenter.Y;

                foreach (var s in shapesToDistribute)
                {
                    totalShapeHeight += s.AbsoluteHeight;
                }

                spaceBetweenShapes = referenceHeight / (shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(1);
                shapesToDistribute.RemoveAt(0);

                shapesToDistribute = Graphics.SortShapesByTop(shapesToDistribute);

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];
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
                    throw new Exception(ErrorMessageFewerThanThreeSelection);
                }

                shapesToDistribute = Graphics.SortShapesByTop(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualTop + shapesToDistribute[0].AbsoluteHeight;
                var bottomMostShape = shapesToDistribute[shapesToDistribute.Count - 1];
                referenceHeight = bottomMostShape.VisualTop + bottomMostShape.AbsoluteHeight - shapesToDistribute[0].VisualTop;

                foreach (var s in shapesToDistribute)
                {
                    totalShapeHeight += s.AbsoluteHeight;
                }
                spaceBetweenShapes = (referenceHeight - totalShapeHeight) / (shapesToDistribute.Count - 1);

                shapesToDistribute.RemoveAt(shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(0);

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];
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
                    throw new Exception(ErrorMessageFewerThanThreeSelection);
                }

                shapesToDistribute = Graphics.SortShapesByTop(selectedShapes);
                startingPoint = shapesToDistribute[0].VisualCenter.Y;
                referenceHeight = shapesToDistribute[shapesToDistribute.Count - 1].VisualCenter.Y - shapesToDistribute[0].VisualCenter.Y;

                foreach (var s in shapesToDistribute)
                {
                    totalShapeHeight += s.AbsoluteHeight;
                }

                spaceBetweenShapes = referenceHeight / (shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(shapesToDistribute.Count - 1);
                shapesToDistribute.RemoveAt(0);

                for (var i = 0; i < shapesToDistribute.Count; i++)
                {
                    var currShape = shapesToDistribute[i];
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

        public static void DistributeGrid(List<PPShape> selectedShapes, int rowLength, int colLength)
        {
            if (DistributeReference == DistributeReferenceObject.ExtremeShapes)
            {
                throw new Exception(ErrorMessageFunctionNotSupportedForExtremeShapes);
            }

            if (DistributeReference == DistributeReferenceObject.Slide)
            {
                throw new Exception(ErrorMessageFunctionNotSupportedForSlide);
            }

            var isFirstTwoShapes = DistributeReference == DistributeReferenceObject.FirstTwoShapes;
            var isFirstShape = DistributeReference == DistributeReferenceObject.FirstShape;
            var isObjectCenter = DistributeSpaceReference == DistributeSpaceReferenceObject.ObjectCenter;
            var isObjectBoundary = DistributeSpaceReference == DistributeSpaceReferenceObject.ObjectBoundary;

            var colLengthGivenFullRows = (int) Math.Ceiling((double) selectedShapes.Count/rowLength);
            if (colLength <= colLengthGivenFullRows)
            {
                if (isFirstTwoShapes && isObjectCenter)
                {
                    var startAnchorCenter = selectedShapes[0].VisualCenter;
                    var endAnchorCenter = selectedShapes[1].VisualCenter;
                    var rowWidth = endAnchorCenter.X - startAnchorCenter.X;
                    var colHeight = endAnchorCenter.Y - startAnchorCenter.Y;
                    DistributeGridByRowWithAnchors(selectedShapes, rowLength, colLength, rowWidth, colHeight);
                }
                else if (isFirstTwoShapes && isObjectBoundary)
                {
                    var startAnchor = selectedShapes[0];
                    var endAnchor = selectedShapes[1];
                    var rowWidth = endAnchor.VisualLeft + endAnchor.AbsoluteWidth - startAnchor.VisualLeft;
                    var colHeight = endAnchor.VisualTop + endAnchor.AbsoluteHeight - startAnchor.VisualTop;
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
                    var startAnchorCenter = selectedShapes[0].VisualCenter;
                    var endAnchorCenter = selectedShapes[1].VisualCenter;
                    var rowWidth = endAnchorCenter.X - startAnchorCenter.X;
                    var colHeight = endAnchorCenter.Y - startAnchorCenter.Y;
                    DistributeGridByColWithAnchors(selectedShapes, rowLength, colLength, rowWidth, colHeight);
                }
                else if (isFirstTwoShapes && isObjectBoundary)
                {
                    var startAnchor = selectedShapes[0];
                    var endAnchor = selectedShapes[1];
                    var rowWidth = endAnchor.VisualLeft + endAnchor.AbsoluteWidth - startAnchor.VisualLeft;
                    var colHeight = endAnchor.VisualTop + endAnchor.AbsoluteHeight - startAnchor.VisualTop;
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
            var refPoint = selectedShapes[0].VisualCenter;

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
                currentShape.IncrementLeft(posX - currentShape.VisualCenter.X);
                currentShape.IncrementTop(posY - currentShape.VisualCenter.Y);

                posX += GetSpaceBetweenShapes(differenceIndex, differenceIndex + 1, rowDifferences, MarginLeft, MarginRight);
                differenceIndex++;
            }
        }

        public static void DistributeGridByCol(List<PPShape> selectedShapes, int rowLength, int colLength)
        {
            var refPoint = selectedShapes[0].VisualCenter;

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
                var center = currentShape.VisualCenter;
                currentShape.IncrementLeft(posX - center.X);
                currentShape.IncrementTop(posY - center.Y);

                posX += GetSpaceBetweenShapes(augmentedShapeIndex%rowLength, augmentedShapeIndex%rowLength + 1, rowDifferences, MarginLeft, MarginRight);
                augmentedShapeIndex++;
            }
        }

        public static void DistributeGridByRowWithAnchors(List<PPShape> selectedShapes, int rowLength, int colLength, float rowWidth, float colHeight)
        {
            if (selectedShapes.Count < 2)
            {
                throw new Exception(ErrorMessageFewerThanTwoSelection);
            }

            var startingAnchor = selectedShapes[0].VisualCenter;
            var endingAnchor = selectedShapes[1].VisualCenter;

            var endAnchor = selectedShapes[1];
            selectedShapes.RemoveAt(1);
            selectedShapes.Add(endAnchor);

            var rowDifference = rowWidth / (rowLength - 1);
            var colDifference = colHeight / (colLength - 1);

            var gridSpaces = new float[selectedShapes.Count, 2];

            for (int i = 0; i < selectedShapes.Count; i++)
            {
                gridSpaces[i, 0] = rowDifference;
                gridSpaces[i, 1] = colDifference;
            }

            DistributeGridByRow(selectedShapes, rowLength, colLength, gridSpaces, 0, selectedShapes.Count - 1);
        }

        public static void DistributeGridByRow(List<PPShape> selectedShapes, int rowLength, int colLength, float[,] gridSpaces, int start, int end)
        {
            var numShapes = selectedShapes.Count;
            var numIndicesToSkip = IndicesToSkip(numShapes, rowLength, DistributeGridAlignment);

            var startingAnchor = selectedShapes[0].VisualCenter;

            var rowDifferences = GetLongestWidthsOfRowsByCol(selectedShapes, rowLength, colLength, numIndicesToSkip);
            var colDifferences = GetLongestHeightsOfColsByCol(selectedShapes, rowLength, colLength, numIndicesToSkip);

            var posX = startingAnchor.X;
            var posY = startingAnchor.Y;
            var remainder = numShapes % rowLength;

            for (var i = start; i < end; i++)
            {
                //Start of new row
                if (i % rowLength == 0 && i != 0)
                {
                    posX = startingAnchor.X;
                    posY += gridSpaces[i, 1];
                }

                //If last row, offset by num of indices to skip
                if (numShapes - i == remainder)
                {
                    posX += numIndicesToSkip * gridSpaces[i, 0];
                }

                var currentShape = selectedShapes[i];
                currentShape.IncrementLeft(posX - currentShape.VisualCenter.X);
                currentShape.IncrementTop(posY - currentShape.VisualCenter.Y);

                posX += gridSpaces[i, 0];
            }
        }

        public static void DistributeGridByRowWithAnchorsByEdge(List<PPShape> selectedShapes, int rowLength, int colLength, float rowWidth, float colHeight)
        {
            if (selectedShapes.Count < 2)
            {
                throw new Exception(ErrorMessageFewerThanTwoSelection);
            }

            var numShapes = selectedShapes.Count;

            var startAnchor = selectedShapes[0];
            var endAnchor = selectedShapes[1];
            selectedShapes.RemoveAt(1);
            selectedShapes.Add(endAnchor);

            var startingAnchor = selectedShapes[0].VisualCenter;

            var longestRow = rowWidth;
            var longestCol = colHeight;

            var colDifferences = GetLongestHeightsOfColsByRow(selectedShapes, rowLength, colLength);

            for (int i = 0; i < colDifferences.Length; i++)
            {
                longestCol -= colDifferences[i];
            }

            var posX = startingAnchor.X;
            var posY = startAnchor.VisualTop + colDifferences[0] / 2;
            var rowDifference = longestRow;
            var colDifference = longestCol / (colDifferences.Length - 1);

            for (var i = 0; i < numShapes - 1; i++)
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

                var currentShape = selectedShapes[i];
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
            var numShapes = selectedShapes.Count;

            var startingAnchor = selectedShapes[0].VisualCenter;

            var posX = startingAnchor.X;
            var posY = startingAnchor.Y;

            var longestRow = GetLongestRowWidthByRow(selectedShapes, rowLength);
            var colDifferences = GetLongestHeightsOfColsByRow(selectedShapes, rowLength, colLength);

            var rowDifference = longestRow;

            for (var i = 0; i < numShapes; i++)
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
                        posY += GetSpaceBetweenShapes(i / rowLength - 1, i / rowLength, colDifferences, MarginTop, MarginBottom);
                    }
                }

                var currentShape = selectedShapes[i];
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
                throw new Exception(ErrorMessageFewerThanTwoSelection);
            }

            var startingAnchor = selectedShapes[0].VisualCenter;
            var endingAnchor = selectedShapes[1].VisualCenter;

            var rowDifference = rowWidth / (rowLength - 1);
            var colDifference = colHeight / (colLength - 1);

            var endAnchor = selectedShapes[1];
            selectedShapes.RemoveAt(1);
            selectedShapes.Add(endAnchor);

            var gridSpaces = new float[selectedShapes.Count, 2];

            for (int i = 0; i < selectedShapes.Count; i++)
            {
                gridSpaces[i, 0] = rowDifference;
                gridSpaces[i, 1] = colDifference;
            }

            DistributeGridByCol(selectedShapes, rowLength, colLength, gridSpaces, 0, selectedShapes.Count - 1);
        }

        public static void DistributeGridByCol(List<PPShape> selectedShapes, int rowLength, int colLength, float[,] gridSpaces, int start, int end)
        {
            var numShapes = selectedShapes.Count;

            var numIndicesToSkip = IndicesToSkip(numShapes, colLength, DistributeGridAlignment);

            var startingAnchor = selectedShapes[0].VisualCenter;

            var posX = startingAnchor.X;
            var posY = startingAnchor.Y;
            var remainder = colLength - (rowLength * colLength - numShapes);
            var augmentedShapeIndex = 0;

            for (var i = start; i < end; i++)
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
                    posY += gridSpaces[i, 1];
                }

                var currentShape = selectedShapes[i];
                var center = currentShape.VisualCenter;
                currentShape.IncrementLeft(posX - center.X);
                currentShape.IncrementTop(posY - center.Y);

                posX += gridSpaces[i, 0];
                augmentedShapeIndex++;
            }
        }

        public static void DistributeGridByColByEdge(List<PPShape> selectedShapes, int rowLength, int colLength)
        {
            var numShapes = selectedShapes.Count;
            var startingAnchor = selectedShapes[0].VisualCenter;

            var posX = startingAnchor.X;
            var posY = startingAnchor.Y;
            var remainder = colLength - (rowLength * colLength - numShapes);
            var augmentedShapeIndex = 0;

            var longestRow = GetLongestRowWidthByCol(selectedShapes, rowLength, colLength);
            var colDifferences = GetLongestHeightsOfColsByCol(selectedShapes, rowLength, colLength, 0);
            var rowDifference = longestRow;

            for (var i = 0; i < numShapes; i++)
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
                        posY += GetSpaceBetweenShapes(augmentedShapeIndex / rowLength - 1, augmentedShapeIndex / rowLength, colDifferences, MarginTop, MarginBottom);
                    }
                }

                var currentShape = selectedShapes[i];
                var center = currentShape.VisualCenter;
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
                throw new Exception(ErrorMessageFewerThanTwoSelection);
            }

            var numShapes = selectedShapes.Count;

            var startAnchor = selectedShapes[0];
            var endAnchor = selectedShapes[1];
            selectedShapes.RemoveAt(1);
            selectedShapes.Add(endAnchor);

            var startingAnchor = selectedShapes[0].VisualCenter;

            var longestRow = rowWidth;
            var longestCol = colHeight;

            var colDifferences = GetLongestHeightsOfColsByCol(selectedShapes, rowLength, colLength, 0);

            for (int i = 0; i < colDifferences.Length; i++)
            {
                longestCol -= colDifferences[i];
            }

            var posX = startingAnchor.X;
            var posY = startAnchor.VisualTop + colDifferences[0] / 2;
            var rowDifference = longestRow;
            var colDifference = longestCol / (colDifferences.Length - 1);
            var remainder = colLength - (rowLength * colLength - numShapes);
            var augmentedShapeIndex = 0;

            for (var i = 0; i < numShapes - 1; i++)
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

                var currentShape = selectedShapes[i];
                var center = currentShape.VisualCenter;
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
            var isAtSecondShape = DistributeRadialReference == DistributeRadialReferenceObject.AtSecondShape;
            var isSecondThirdShape = DistributeRadialReference == DistributeRadialReferenceObject.SecondThirdShape;
            var isObjectBoundary = DistributeSpaceReference == DistributeSpaceReferenceObject.ObjectBoundary;
            var isObjectCenter = DistributeSpaceReference == DistributeSpaceReferenceObject.ObjectCenter;
            
            Drawing.PointF origin;
            float referenceAngle, startingAngle;

            if (isAtSecondShape && isObjectBoundary)
            {
                if (selectedShapes.Count < 3)
                {
                    throw new Exception(ErrorMessageFewerThanThreeSelection);
                }

                origin = Graphics.GetCenterPoint(selectedShapes[1]);

                var boundaryAngles = GetShapeBoundaryAngles(origin, selectedShapes[2]);
                startingAngle = boundaryAngles[1];
                var endingAngle = boundaryAngles[0];

                if (startingAngle == 0 && boundaryAngles[1] == 360)
                {
                    throw new Exception(ErrorMessageFunctionNotSuppertedForOverlapRefShapeCenter);
                }

                referenceAngle = endingAngle - startingAngle;
                if (referenceAngle < 0)
                {
                    referenceAngle += 360;
                }

                var offset = endingAngle - startingAngle;
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
                    throw new Exception(ErrorMessageFewerThanThreeSelection);
                }
                
                origin = Graphics.GetCenterPoint(selectedShapes[1]);
                
                startingAngle = (float)AngleBetweenTwoPoints(origin, GetVisualCenter(selectedShapes[2]));
                referenceAngle = 360;

                DistributeShapesWithinAngleForCenter(selectedShapes, origin, startingAngle, referenceAngle, 3);
            }
            else if (isSecondThirdShape && isObjectBoundary)
            {
                if (selectedShapes.Count < 4)
                {
                    throw new Exception(ErrorMessageFewerThanFourSelection);
                }

                origin = Graphics.GetCenterPoint(selectedShapes[1]);
                var startingShapeBoundaryAngles = GetShapeBoundaryAngles(origin, selectedShapes[2]);
                var endingShapeBoundaryAngles = GetShapeBoundaryAngles(origin, selectedShapes[3]);
                startingAngle = startingShapeBoundaryAngles[0];
                var endingAngle = endingShapeBoundaryAngles[1];

                if ((startingAngle == 0 && startingShapeBoundaryAngles[1] == 360)
                    || (endingShapeBoundaryAngles[0] == 0 && endingAngle == 360))
                {
                    throw new Exception(ErrorMessageFunctionNotSuppertedForOverlapRefShapeCenter);
                }

                var startingShapeAngle = startingShapeBoundaryAngles[1] - startingAngle;
                if (startingShapeAngle < 0)
                {
                    startingShapeAngle += 360;
                }

                var endingShapeAngle = endingAngle - endingShapeBoundaryAngles[0];
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
                    throw new Exception(ErrorMessageFewerThanFourSelection);
                }

                origin = Graphics.GetCenterPoint(selectedShapes[1]);
                startingAngle = (float)AngleBetweenTwoPoints(origin, GetVisualCenter(selectedShapes[2]));
                var endingAngle = (float)AngleBetweenTwoPoints(origin, GetVisualCenter(selectedShapes[3]));

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
            var shapeAngleInfos = new List<ShapeAngleInfo>();

            for (int i = startingIndex; i <= selectedShapes.Count; i++)
            {
                var angle = (float)AngleBetweenTwoPoints(origin, Graphics.GetCenterPoint(selectedShapes[i]));
                var angleFromStart = (angle + (360 - startingAngle)) % 360;

                var shapeAngleInfo = new ShapeAngleInfo(selectedShapes[i], angleFromStart);
                shapeAngleInfos.Add(shapeAngleInfo);
            }

            shapeAngleInfos = shapeAngleInfos.OrderBy(x => x.Angle).ToList();

            var angleBetweenShapes = referenceAngle / (shapeAngleInfos.Count + 1);
            var endingAngle = 0f;

            foreach (var shapeAngleInfo in shapeAngleInfos)
            {
                endingAngle += angleBetweenShapes;

                var rotationAngle = endingAngle - shapeAngleInfo.Angle;
                Rotate(shapeAngleInfo.Shape, origin, rotationAngle, DistributeShapeOrientation);
            }
        }

        public static void DistributeShapesWithinAngleForBoundary(ShapeRange selectedShapes, Drawing.PointF origin, float startingAngle,
             float referenceAngle, int startingIndex, float startingShapeAngle = 0, float endingShapeAngle = 0, float offset = 0)
        {
            var shapeAngleInfos = new List<ShapeAngleInfo>();

            var boundaryShapeAngle = startingShapeAngle + endingShapeAngle;
            if (boundaryShapeAngle >= referenceAngle)
            {
                boundaryShapeAngle = 0;
            }

            var count = 0;
            var isStable = false;
            while (!isStable && count < 20)
            {
                var totalShapeAngle = boundaryShapeAngle;

                if (count == 0)
                {
                    for (int i = startingIndex; i <= selectedShapes.Count; i++)
                    {
                        float shapeAngle;
                        var angle = GetShapeAngleInfo(selectedShapes[i], origin, startingAngle, totalShapeAngle, out shapeAngle);
                        totalShapeAngle += shapeAngle;

                        var shapeAngleInfo = new ShapeAngleInfo(selectedShapes[i], angle, shapeAngle);
                        shapeAngleInfos.Add(shapeAngleInfo);
                    }

                    shapeAngleInfos = shapeAngleInfos.OrderBy(x => (x.Angle - offset) % 360).ToList();
                }
                else
                {
                    foreach (var shapeAngleInfo in shapeAngleInfos)
                    {
                        float shapeAngle;
                        var angle = GetShapeAngleInfo(shapeAngleInfo.Shape, origin, startingAngle, totalShapeAngle, out shapeAngle);
                        totalShapeAngle += shapeAngle;

                        shapeAngleInfo.Angle = angle;
                        shapeAngleInfo.ShapeAngle = shapeAngle;
                    }
                }

                var angleBetweenShapes = (referenceAngle - totalShapeAngle) / (shapeAngleInfos.Count + 1);
                var endingAngle = (boundaryShapeAngle == 0) ? angleBetweenShapes : startingShapeAngle + angleBetweenShapes;

                isStable = true;

                foreach (var shapeAngleInfo in shapeAngleInfos)
                {
                    var rotationAngle = (endingAngle - shapeAngleInfo.Angle) % 360;
                    if (rotationAngle > threshold || rotationAngle < -threshold)
                    {
                        isStable = false;
                        Rotate(shapeAngleInfo.Shape, origin, rotationAngle, DistributeShapeOrientation);
                    }

                    endingAngle += shapeAngleInfo.ShapeAngle + angleBetweenShapes;
                }

                count++;
            }
        }

        private static float GetShapeAngleInfo(Shape shape, Drawing.PointF origin, float startingAngle, float totalShapeAngle,
            out float shapeAngle)
        {
            var boundaryAnglesFromStart = GetShapeBoundaryAngles(origin, shape);
            if (boundaryAnglesFromStart[0] == 0 && boundaryAnglesFromStart[1] == 360)
            {
                throw new Exception(ErrorMessageFunctionNotSuppertedForOverlapRefShapeCenter);
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
                throw new Exception(ErrorMessageFewerThanTwoSelection);
            }

            var sortedShapes = selectedShapes;

            if (!IsSwapByClickOrder)
            {
                if (ListIsPreviouslySelected(selectedShapes, prevSelectedShapes))
                {
                    sortedShapes = SortPPShapesByName(selectedShapes, prevSortedShapeNames);
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

            var shapeNames = new List<string>();

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

                if (i != 0 && !IsSwapByClickOrder && !isPreview)
                {
                    shapeNames.Add(currentShape.Name);
                }
            }

            if (!IsSwapByClickOrder && !isPreview)
            {
                shapeNames.Insert(0, sortedShapes[0].Name);
                prevSortedShapeNames = shapeNames;
                SaveSelectedList(selectedShapes, prevSelectedShapes);
            }
        }

        #endregion

        #region Adjustment
        
        public static void Rotate(Shape shape, Drawing.PointF origin, float angle, RadialShapeOrientationObject shapeOrientation)
        {
            var unrotatedCenter = Graphics.GetCenterPoint(shape);
            var rotatedCenter = Graphics.RotatePoint(unrotatedCenter, origin, angle);

            shape.Left += (rotatedCenter.X - unrotatedCenter.X);
            shape.Top += (rotatedCenter.Y - unrotatedCenter.Y);

            if (shapeOrientation == RadialShapeOrientationObject.Dynamic)
            {
                shape.Rotation = AddAngles(shape.Rotation, angle);
            }
        }

        #endregion

        #region Snap

        public static void SnapVertical(IList<Shape> selectedShapes)
        {
            foreach (var s in selectedShapes)
            {
                SnapShapeVertical(s);
            }
        }

        public static void SnapHorizontal(IList<Shape> selectedShapes)
        {
            foreach (var s in selectedShapes)
            {
                SnapShapeHorizontal(s);
            }
        }

        public static void SnapAway(IList<Shape> shapes)
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
            }

            if (!isAllSameDir || lastDir == None || lastDir == Up)
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
            var shapeIsVertical = shape.Height > shape.Width;

            if (NearlyEqual(shape.Height, shape.Width, Epsilon))
            {
                float defaultUpAngle = 0;
                var hasDefaultDirection = shapeDefaultUpAngle.TryGetValue(shape.AutoShapeType, out defaultUpAngle);
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
                throw new Exception(ErrorMessageNoSelection);
            }

            for (var i = 1; i <= selectedShapes.Count; i++)
            {
                var currentShape = selectedShapes[i];
                var currentRotation = currentShape.Rotation;
                currentShape.Flip(MsoFlipCmd.msoFlipHorizontal);
                currentShape.Rotation = currentRotation;
            }
        }

        public static void FlipVertical(ShapeRange selectedShapes)
        {
            if (selectedShapes.Count < 1)
            {
                throw new Exception(ErrorMessageNoSelection);
            }

            for (var i = 1; i <= selectedShapes.Count; i++)
            {
                var currentShape = selectedShapes[i];
                var currentRotation = currentShape.Rotation;
                currentShape.Flip(MsoFlipCmd.msoFlipVertical);
                currentShape.Rotation = currentRotation;
            }
        }

        #endregion

        #endregion

        #region Util

        public static double AngleBetweenTwoPoints(Drawing.PointF refPoint, Drawing.PointF pt)
        {
            var angle = Math.Atan((pt.Y - refPoint.Y)/(pt.X - refPoint.X))*180/Math.PI;

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
            var distance = Math.Sqrt(Math.Pow(pt.X - refPoint.X, 2) + Math.Pow(refPoint.Y - pt.Y, 2));
            return distance;
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

            if (!NearlyEqual(phaseInFloat, (float) Math.Round(phaseInFloat), Epsilon))
            {
                return None;
            }

            var phase = (int) Math.Round(phaseInFloat);

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

        public static float GetLongestRowWidthByRow(List<PPShape> shapes, int rowLength)
        {
            var longestRow = 0.0f;
            var longestRowSoFar = 0.0f;
            var numShapes = shapes.Count;
            var remainder = numShapes % rowLength;

            for (var i = 0; i < numShapes; i++)
            {
                var rowIndex = i % rowLength;

                if (rowIndex == 0)
                {
                    if (longestRowSoFar > longestRow)
                    {
                        longestRow = longestRowSoFar;
                    }
                    longestRowSoFar = -(MarginLeft + MarginRight);
                }
                longestRowSoFar += (shapes[i].AbsoluteWidth + MarginLeft + MarginRight);
            }

            if (longestRowSoFar > longestRow)
            {
                longestRow = longestRowSoFar;
            }

            return longestRow;
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

        public static float GetLongestRowWidthByCol(List<PPShape> shapes, int rowLength, int colLength)
        {
            float longestWidth = 0;
            var numShapes = shapes.Count;
            
            var augmentedShapeIndex = 0;
            var remainder = colLength - (rowLength * colLength - numShapes);
            float rowSoFar = 0;

            for (var i = 0; i < numShapes; i++)
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
                        rowSoFar += (shapes[i].AbsoluteWidth + MarginLeft + MarginRight);
                        augmentedShapeIndex++;
                        i++;
                    }

                    if (rowSoFar > longestWidth)
                    {
                        longestWidth = rowSoFar;
                    }
                    rowSoFar = -(MarginLeft + MarginRight);
                }

                if (i < numShapes)
                {
                    rowSoFar += (shapes[i].AbsoluteWidth + MarginLeft + MarginRight);
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

            if (alignment == GridAlignment.AlignLeft || alignment == GridAlignment.None || numOfShapesInLastRow == 0)
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

        private static float[] GetShapeBoundaryAngles(Drawing.PointF origin, Shape shape)
        {
            var ppShape = new PPShape(shape, false);
            var points = ppShape.Points;
            var pointAngles = new List<float>();

            foreach (var point in points)
            {
                var angle = (float)AngleBetweenTwoPoints(origin, point);
                pointAngles.Add(angle);
            }

            var isSpanAcross0Degrees = false;
            var hasTurningPoint = false;
            bool isCurrentClockwise;
            var isPreviousClockwise = pointAngles[0] - pointAngles[pointAngles.Count - 1] >= 0;
            var turningPointAngles = new List<float>();

            var boundaryAngles = new float[2];
            boundaryAngles[0] = pointAngles[0];
            boundaryAngles[1] = pointAngles[0];

            for (int i = 1; i < pointAngles.Count; i++)
            {
                var previousAngle = pointAngles[i - 1];
                var currentAngle = pointAngles[i];
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
            var duplicateShape = shape.Duplicate()[1];
            duplicateShape.Left = shape.Left;
            duplicateShape.Top = shape.Top;

            var duplicatePPShape = new PPShape(duplicateShape);
            var visualCenter = duplicatePPShape.VisualCenter;
            duplicateShape.Delete();

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
                    var shapePos = selectedShapes[i].VisualCenter;
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
                var shapePos = selectedShapes[i].VisualCenter;
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
            var selectedShapes = new List<PPShape>();
            for (var i = 1; i <= toAlign.Count; i++)
            {
                var s = toAlign[i];
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

        private static Drawing.PointF GetSwapReferencePoint(PPShape shape, SwapReference r)
        {
            switch (r)
            {
                case SwapReference.TopLeft:
                    return shape.VisualTopLeft;
                case SwapReference.TopCenter:
                    return shape.VisualTopCenter;
                case SwapReference.TopRight:
                    return shape.VisualTopRight;
                case SwapReference.MiddleLeft:
                    return shape.VisualMiddleLeft;
                case SwapReference.MiddleCenter:
                    return shape.VisualCenter;
                case SwapReference.MiddleRight:
                    return shape.VisualMiddleRight;
                case SwapReference.BottomLeft:
                    return shape.VisualBottomLeft;
                case SwapReference.BottomCenter:
                    return shape.VisualBottomCenter;
                case SwapReference.BottomRight:
                    return shape.VisualBottomRight;
                default:
                    return shape.VisualCenter;
            }
        }

        private static void InitDefaultAlignSettings()
        {
            AlignReferToShape();
        }
        private static void InitDefaultAdjoinSettings()
        {
            AdjoinWithAligning();
        }

        private static void InitDefaultDistributeSettings()
        {
            MarginTop = 5;
            MarginBottom = 5;
            MarginLeft = 5;
            MarginRight = 5;
            DistributeGridAlignment = GridAlignment.AlignLeft;
            DistributeReferToFirstTwoShapes();
            DistributeReferToSecondThirdShape();
            DistributeSpaceByBoundaries();
            DistributeShapeOrientationToFixed();
        }

        private static void InitDefaultSwapSettings()
        {
            IsSwapByClickOrder = false;
            SwapReferencePoint = SwapReference.MiddleCenter;
            prevSelectedShapes = new Dictionary<string, Drawing.PointF>();
            prevSortedShapeNames = new List<string>();
        }
        
        private static void InitDefaultReorientSettings()
        {
            ReorientShapeOrientationToFixed();
        }

        public static void InitPositionsLab()
        {
            InitDefaultAlignSettings();
            InitDefaultAdjoinSettings();
            InitDefaultDistributeSettings();
            InitDefaultSwapSettings();
            InitDefaultReorientSettings();
            InitDefaultShapesAngles();
        }

        #endregion
    }
}
