using System;
using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PositionsLab;
using PowerPointLabs.Utils;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using SlideUtil = Test.Util.SlideUtil;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class PositionsLabSwapTest : BasePositionsLabTest
    {
        private List<string> _allShapes;
        private List<string> _swapShapes;

        private const int OriginalShapesSlideNo = 3;
        private const string Square = "Square";
        private const string Arrow = "Arrow";
        private const string Line = "Line";
        private const string Circle = "Circle";
        private const string Triangle = "Triangle";

        //Results of Operations
        private const int SwapLeftToRight1Slide = 5;
        private const int SwapLeftToRight2Slide = 6;
        private const int SwapLeftToRight3Slide = 7;
        private const int SwapLeftToRight4Slide = 8;

        private const int SwapClick1Slide = 10;
        private const int SwapClick2Slide = 11;
        private const int SwapClick3Slide = 12;
        private const int SwapClick4Slide = 13;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabSwap.pptx";
        }

        [TestInitialize]
        public void TestInitialize()
        {
            PositionsLabMain.InitPositionsLab();

            _allShapes = new List<string> { Square, Arrow, Line, Circle, Triangle };
            _swapShapes = new List<string> { Circle, Arrow, Square, Triangle };
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSwapLeftToRight()
        {
            PositionsLabSettings.IsSwapByClickOrder = false;
            PositionsLabSettings.SwapReferencePoint = PositionsLabSettings.SwapReference.MiddleCenter;

            PowerPoint.ShapeRange shapesToSwap = GetShapes(OriginalShapesSlideNo, _swapShapes);

            Action<List<PPShape>, bool> positionsAction = (shapes, isPreview) => PositionsLabMain.Swap(shapes, isPreview);

            ExecutePositionsAction(positionsAction, shapesToSwap, false);
            PowerPoint.ShapeRange expectedShapes1Swap = GetShapes(SwapLeftToRight1Slide, _allShapes);
            PowerPoint.ShapeRange actualShapes1Swap = GetShapes(OriginalShapesSlideNo, _allShapes);
            CheckShapes(expectedShapes1Swap, actualShapes1Swap);

            ExecutePositionsAction(positionsAction, shapesToSwap, false);
            PowerPoint.ShapeRange expectedShapes2Swap = GetShapes(SwapLeftToRight2Slide, _allShapes);
            PowerPoint.ShapeRange actualShapes2Swap = GetShapes(OriginalShapesSlideNo, _allShapes);
            CheckShapes(expectedShapes2Swap, actualShapes2Swap);

            ExecutePositionsAction(positionsAction, shapesToSwap, false);
            PowerPoint.ShapeRange expectedShapes3Swap = GetShapes(SwapLeftToRight3Slide, _allShapes);
            PowerPoint.ShapeRange actualShapes3Swap = GetShapes(OriginalShapesSlideNo, _allShapes);
            CheckShapes(expectedShapes3Swap, actualShapes3Swap);

            ExecutePositionsAction(positionsAction, shapesToSwap, false);
            PowerPoint.ShapeRange expectedShapes4Swap = GetShapes(SwapLeftToRight4Slide, _allShapes);
            PowerPoint.ShapeRange actualShapes4Swap = GetShapes(OriginalShapesSlideNo, _allShapes);
            CheckShapes(expectedShapes4Swap, actualShapes4Swap);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSwapClickOrder()
        {
            PositionsLabSettings.IsSwapByClickOrder = true;
            PositionsLabSettings.SwapReferencePoint = PositionsLabSettings.SwapReference.MiddleCenter;

            PowerPoint.ShapeRange shapesToSwap = GetShapes(OriginalShapesSlideNo, _swapShapes);

            Action<List<PPShape>, bool> positionsAction = (shapes, isPreview) => PositionsLabMain.Swap(shapes, isPreview);

            ExecutePositionsAction(positionsAction, shapesToSwap, false);
            PowerPoint.ShapeRange expectedShapes1Swap = GetShapes(SwapClick1Slide, _allShapes);
            PowerPoint.ShapeRange actualShapes1Swap = GetShapes(OriginalShapesSlideNo, _allShapes);
            CheckShapes(expectedShapes1Swap, actualShapes1Swap);
            
            ExecutePositionsAction(positionsAction, shapesToSwap, false);
            PowerPoint.ShapeRange expectedShapes2Swap = GetShapes(SwapClick2Slide, _allShapes);
            PowerPoint.ShapeRange actualShapes2Swap = GetShapes(OriginalShapesSlideNo, _allShapes);
            CheckShapes(expectedShapes2Swap, actualShapes2Swap);

            ExecutePositionsAction(positionsAction, shapesToSwap, false);
            PowerPoint.ShapeRange expectedShapes3Swap = GetShapes(SwapClick3Slide, _allShapes);
            PowerPoint.ShapeRange actualShapes3Swap = GetShapes(OriginalShapesSlideNo, _allShapes);
            CheckShapes(expectedShapes3Swap, actualShapes3Swap);

            ExecutePositionsAction(positionsAction, shapesToSwap, false);
            PowerPoint.ShapeRange expectedShapes4Swap = GetShapes(SwapClick4Slide, _allShapes);
            PowerPoint.ShapeRange actualShapes4Swap = GetShapes(OriginalShapesSlideNo, _allShapes);
            CheckShapes(expectedShapes4Swap, actualShapes4Swap);
        }

        private new void CheckShapes(PowerPoint.ShapeRange expectedShapes, PowerPoint.ShapeRange actualShapes)
        {
            foreach (PowerPoint.Shape actualShape in actualShapes)
            {
                bool isFound = false;

                foreach (PowerPoint.Shape expectedShape in expectedShapes)
                {
                    if (!actualShape.Name.Equals(expectedShape.Name)) continue;
                    isFound = true;
                    SlideUtil.IsSameShape(expectedShape, actualShape);
                    SlideUtil.IsSameZOrderPosition(expectedShape, actualShape);
                    break;
                }

                if (!isFound)
                {
                    Assert.Fail("Unable to find corresponding actual shape");
                }
            }
        }
    }
}
