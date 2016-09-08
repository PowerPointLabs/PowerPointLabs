using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointLabs.PositionsLab;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using PowerPointLabs.Utils;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class PositionsLabFlipTest : BasePositionsLabTest
    {
        private List<string> _shapeNames;

        private const int OriginalShapes0DegreesSlideNo = 3;
        private const int OriginalShapes315DegreesSlideNo = 4;
        private const int OriginalShapes225DegreesSlideNo = 5;
        private const int OriginalShapes135DegreesSlideNo = 6;
        private const int OriginalShapes45DegreesSlideNo = 7;

        private const string Rectangle3 = "Rectangle 31";
        private const string Oval4 = "Oval 41";
        private const string Rectangle6 = "Rectangle 61";
        private const string Picture3 = "Picture 31";
        private const string Picture4 = "Picture 41";
        private const string Picture2 = "Picture 21";
        private const string LeftArrow2 = "Left Arrow 21";
        private const string UpArrow10 = "Up Arrow 101";
        private const string DownArrow11 = "Down Arrow 111";
        private const string LeftRightArrow12 = "Left-Right Arrow 121";
        private const string UpDownArrow13 = "Up-Down Arrow 131";
        private const string QuadArrow14 = "Quad Arrow 141";
        private const string LeftRightUpArrow15 = "Left-Right-Up Arrow 151";
        private const string BentArrow16 = "Bent Arrow 161";
        private const string UTurnArrow17 = "U-Turn Arrow 171";
        private const string LeftUpArrow18 = "Left-Up Arrow 181";
        private const string BentUpArrow19 = "Bent-Up Arrow 191";
        private const string CurvedRightArrow20 = "Curved Right Arrow 201";
        private const string CurvedLeftArrow21 = "Curved Left Arrow 211";
        private const string CurvedDownArrow22 = "Curved Down Arrow 221";
        private const string CurvedUpArrow23 = "Curved Up Arrow 231";
        private const string StripedRightArrow24 = "Striped Right Arrow 241";
        private const string NotchedRightArrow25 = "Notched Right Arrow 251";
        private const string Pentagon26 = "Pentagon 261";
        private const string Chevron27 = "Chevron 271";
        private const string RightArrowCallout28 = "Right Arrow Callout 281";
        private const string DownArrowCallout29 = "Down Arrow Callout 291";
        private const string LeftArrowCallout30 = "Left Arrow Callout 301";
        private const string UpArrowCallout31 = "Up Arrow Callout 311";
        private const string LeftRightArrowCallout32 = "Left-Right Arrow Callout 321";
        private const string QuadArrowCallout33 = "Quad Arrow Callout 331";
        private const string CircularArrow34 = "Circular Arrow 341";
        private const string RightArrow1 = "Right Arrow 11";

        //Results of Operations
        private const int FlipHorizontal0DegreesSlideNo = 9;
        private const int FlipHorizontal315DegreesSlideNo = 10;
        private const int FlipHorizontal225DegreesSlideNo = 11;
        private const int FlipHorizontal135DegreesSlideNo = 12;
        private const int FlipHorizontal45DegreesSlideNo = 13;

        private const int FlipVertical0DegreesSlideNo = 15;
        private const int FlipVertical315DegreesSlideNo = 16;
        private const int FlipVertical225DegreesSlideNo = 17;
        private const int FlipVertical135DegreesSlideNo = 18;
        private const int FlipVertical45DegreesSlideNo = 19;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabFlip.pptx";
        }

        [TestInitialize]
        public void TestInitialize()
        {
            PositionsLabMain.InitPositionsLab();
            _shapeNames = new List<string> { Rectangle3, Oval4, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapes0DegreesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipHorizontal0Degrees()
        {
            InitOriginalShapes(OriginalShapes0DegreesSlideNo, _shapeNames);
            _shapeNames = new List<string> { Rectangle3, Oval4, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };

            var actualShapes = GetShapes(OriginalShapes0DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipHorizontal0DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
            //RestoreShapes(OriginalShapes0DegreesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipHorizontal315Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes315DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipHorizontal315DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipHorizontal225Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes225DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipHorizontal225DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipHorizontal135Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes135DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipHorizontal135DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipHorizontal45Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes45DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipHorizontal45DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipVertical0Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes0DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipVertical0DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipVertical315Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes315DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipVertical315DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipVertical225Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes225DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipVertical225DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipVertical135Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes135DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipVertical135DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFlipVertical45Degrees()
        {
            var actualShapes = GetShapes(OriginalShapes45DegreesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.FlipHorizontal(shapes);
            ExecuteFlipAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(FlipVertical45DegreesSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
