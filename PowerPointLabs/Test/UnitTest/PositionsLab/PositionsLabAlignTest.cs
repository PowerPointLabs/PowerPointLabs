using System;
using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PositionsLab;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class PositionsLabAlignTest : BasePositionsLabTest
    {
        private List<string> _shapeNames;

        private const int OriginalShapesSlideNo = 3;
        private const string UnrotatedRectangle = "Rectangle 3";
        private const string Oval = "Oval 4";
        private const string RotatedArrow = "Right Arrow 5";
        private const string RotatedRectangle = "Rectangle 6";

        //Results of Operations
        private const int AlignShapesLeftToSlideNo = 5;
        private const int AlignShapesRightToSlideNo = 6;
        private const int AlignShapesTopToSlideNo = 7;
        private const int AlignShapesBottomToSlideNo = 8;
        private const int AlignShapesHorizontalToSlideNo = 9;
        private const int AlignShapesVerticalToSlideNo = 10;
        private const int AlignShapesCenterToSlideNo = 11;

        private const int AlignShapesLeftToRefShapeNo = 13;
        private const int AlignShapesRightToRefShapeNo = 14;
        private const int AlignShapesTopToRefShapeNo = 15;
        private const int AlignShapesBottomToRefShapeNo = 16;
        private const int AlignShapesHorizontalToRefShapeNo = 17;
        private const int AlignShapesVerticalToRefShapeNo = 18;
        private const int AlignShapesCenterToRefShapeNo = 19;

        private const int AlignOneShapeLeftDefaultNo = 21;
        private const int AlignOneShapeRightDefaultNo = 22;
        private const int AlignOneShapeTopDefaultNo = 23;
        private const int AlignOneShapeBottomDefaultNo = 24;
        private const int AlignOneShapeHorizontalDefaultNo = 25;
        private const int AlignOneShapeVerticalDefaultNo = 26;
        private const int AlignOneShapeCenterDefaultNo = 27;

        private const int AlignShapesLeftDefaultNo = 28;
        private const int AlignShapesRightDefaultNo = 29;
        private const int AlignShapesTopDefaultNo = 30;
        private const int AlignShapesBottomDefaultNo = 31;
        private const int AlignShapesHorizontalDefaultNo = 32;
        private const int AlignShapesVerticalDefaultNo = 33;
        private const int AlignShapesCenterDefaultNo = 34;

        private const int AlignShapesRadialNo = 36;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabAlign.pptx";
        }

        [TestInitialize]
        public void TestInitialize()
        {
            PositionsLabMain.InitPositionsLab();

            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignLeftToSlide()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.Slide;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = shapes => PositionsLabMain.AlignLeft(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(AlignShapesLeftToSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignRightToSlide()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.Slide;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideWidth = Pres.PageSetup.SlideWidth;

            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, width) => PositionsLabMain.AlignRight(shapes, width);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth);

            PpOperations.SelectSlide(AlignShapesRightToSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignTopToSlide()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.Slide;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = shapes => PositionsLabMain.AlignTop(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(AlignShapesTopToSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignBottomToSlide()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.Slide;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideHeight = Pres.PageSetup.SlideHeight;

            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, height) => PositionsLabMain.AlignBottom(shapes, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideHeight);

            PpOperations.SelectSlide(AlignShapesBottomToSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignHorizontalToSlide()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.Slide;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideHeight = Pres.PageSetup.SlideHeight;

            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, height) => PositionsLabMain.AlignHorizontalCenter(shapes, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideHeight);

            PpOperations.SelectSlide(AlignShapesHorizontalToSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignVerticalToSlide()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.Slide;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideWidth = Pres.PageSetup.SlideWidth;

            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, width) => PositionsLabMain.AlignVerticalCenter(shapes, width);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth);

            PpOperations.SelectSlide(AlignShapesVerticalToSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignCenterToSlide()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.Slide;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideHeight = Pres.PageSetup.SlideHeight;
            float slideWidth = Pres.PageSetup.SlideWidth;

            Action<PowerPoint.ShapeRange, float, float> positionsAction = (shapes, height, width) => PositionsLabMain.AlignCenter(shapes, height, width);
            ExecutePositionsAction(positionsAction, actualShapes, slideHeight, slideWidth);

            PpOperations.SelectSlide(AlignShapesCenterToSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignLeftToRefShape()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.SelectedShape;
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, Oval, RotatedArrow };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = shapes => PositionsLabMain.AlignLeft(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(AlignShapesLeftToRefShapeNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignRightToRefShape()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.SelectedShape;
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, Oval, RotatedArrow };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideWidth = Pres.PageSetup.SlideWidth;

            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, width) => PositionsLabMain.AlignRight(shapes, width);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth);

            PpOperations.SelectSlide(AlignShapesRightToRefShapeNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignTopToRefShape()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.SelectedShape;
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, Oval, RotatedArrow };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = shapes => PositionsLabMain.AlignTop(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(AlignShapesTopToRefShapeNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignBottomToRefShape()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.SelectedShape;
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, Oval, RotatedArrow };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideHeight = Pres.PageSetup.SlideHeight;

            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, height) => PositionsLabMain.AlignBottom(shapes, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideHeight);

            PpOperations.SelectSlide(AlignShapesBottomToRefShapeNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignHorizontalToRefShape()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.SelectedShape;
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, Oval, RotatedArrow };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideHeight = Pres.PageSetup.SlideHeight;

            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, height) => PositionsLabMain.AlignHorizontalCenter(shapes, height);
            ExecutePositionsAction(positionsAction, actualShapes, slideHeight);

            PpOperations.SelectSlide(AlignShapesHorizontalToRefShapeNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignVerticalToRefShape()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.SelectedShape;
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, Oval, RotatedArrow };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideWidth = Pres.PageSetup.SlideWidth;

            Action<PowerPoint.ShapeRange, float> positionsAction = (shapes, width) => PositionsLabMain.AlignVerticalCenter(shapes, width);
            ExecutePositionsAction(positionsAction, actualShapes, slideWidth);

            PpOperations.SelectSlide(AlignShapesVerticalToRefShapeNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignCenterToRefShape()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.SelectedShape;
            _shapeNames = new List<string> { RotatedRectangle, UnrotatedRectangle, Oval, RotatedArrow };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideHeight = Pres.PageSetup.SlideHeight;
            float slideWidth = Pres.PageSetup.SlideWidth;

            Action<PowerPoint.ShapeRange, float, float> positionsAction = (shapes, height, width) => PositionsLabMain.AlignCenter(shapes, height, width);
            ExecutePositionsAction(positionsAction, actualShapes, slideHeight, slideWidth);

            PpOperations.SelectSlide(AlignShapesCenterToRefShapeNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignOneLeftDefault()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.PowerpointDefaults;
            _shapeNames = new List<string> { RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            PositionsLabMain.AlignLeft(actualShapes);

            PpOperations.SelectSlide(AlignOneShapeLeftDefaultNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignOneRightDefault()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.PowerpointDefaults;
            _shapeNames = new List<string> { RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideWidth = Pres.PageSetup.SlideWidth;
            PositionsLabMain.AlignRight(actualShapes, slideWidth);

            PpOperations.SelectSlide(AlignOneShapeRightDefaultNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignOneTopDefault()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.PowerpointDefaults;
            _shapeNames = new List<string> { RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            PositionsLabMain.AlignTop(actualShapes);

            PpOperations.SelectSlide(AlignOneShapeTopDefaultNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignOneBottomDefault()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.PowerpointDefaults;
            _shapeNames = new List<string> { RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideHeight = Pres.PageSetup.SlideHeight;
            PositionsLabMain.AlignBottom(actualShapes, slideHeight);

            PpOperations.SelectSlide(AlignOneShapeBottomDefaultNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignOneHorizontalDefault()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.PowerpointDefaults;
            _shapeNames = new List<string> { RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideHeight = Pres.PageSetup.SlideHeight;
            PositionsLabMain.AlignHorizontalCenter(actualShapes, slideHeight);

            PpOperations.SelectSlide(AlignOneShapeHorizontalDefaultNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignOneVerticalDefault()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.PowerpointDefaults;
            _shapeNames = new List<string> { RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideWidth = Pres.PageSetup.SlideWidth;
            PositionsLabMain.AlignVerticalCenter(actualShapes, slideWidth);

            PpOperations.SelectSlide(AlignOneShapeVerticalDefaultNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignOneCenterDefault()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.PowerpointDefaults;
            _shapeNames = new List<string> { RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideHeight = Pres.PageSetup.SlideHeight;
            float slideWidth = Pres.PageSetup.SlideWidth;
            PositionsLabMain.AlignCenter(actualShapes, slideHeight, slideWidth);

            PpOperations.SelectSlide(AlignOneShapeCenterDefaultNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignOneLeftSingleShape()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.SelectedShape;
            _shapeNames = new List<string> { RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            PositionsLabMain.AlignLeft(actualShapes);

            PpOperations.SelectSlide(AlignOneShapeLeftDefaultNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignOneRightSingleShape()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.SelectedShape;
            _shapeNames = new List<string> { RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            PositionsLabMain.AlignRight(actualShapes, slideWidth);

            PpOperations.SelectSlide(AlignOneShapeRightDefaultNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignOneTopSingleShape()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.SelectedShape;
            _shapeNames = new List<string> { RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            PositionsLabMain.AlignTop(actualShapes);

            PpOperations.SelectSlide(AlignOneShapeTopDefaultNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignOneBottomSingleShape()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.SelectedShape;
            _shapeNames = new List<string> { RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideHeight = Pres.PageSetup.SlideHeight;
            PositionsLabMain.AlignBottom(actualShapes, slideHeight);

            PpOperations.SelectSlide(AlignOneShapeBottomDefaultNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignOneHorizontalSingleShape()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.SelectedShape;
            _shapeNames = new List<string> { RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideHeight = Pres.PageSetup.SlideHeight;
            PositionsLabMain.AlignHorizontalCenter(actualShapes, slideHeight);

            PpOperations.SelectSlide(AlignOneShapeHorizontalDefaultNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignOneVerticalSingleShape()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.SelectedShape;
            _shapeNames = new List<string> { RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            PositionsLabMain.AlignVerticalCenter(actualShapes, slideWidth);

            PpOperations.SelectSlide(AlignOneShapeVerticalDefaultNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignOneCenterSingleShape()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.SelectedShape;
            _shapeNames = new List<string> { RotatedRectangle };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideHeight = Pres.PageSetup.SlideHeight;
            var slideWidth = Pres.PageSetup.SlideWidth;
            PositionsLabMain.AlignCenter(actualShapes, slideHeight, slideWidth);

            PpOperations.SelectSlide(AlignOneShapeCenterDefaultNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignLeftDefault()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.PowerpointDefaults;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            PositionsLabMain.AlignLeft(actualShapes);

            PpOperations.SelectSlide(AlignShapesLeftDefaultNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignRightDefault()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.PowerpointDefaults;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideWidth = Pres.PageSetup.SlideWidth;
            PositionsLabMain.AlignRight(actualShapes, slideWidth);

            PpOperations.SelectSlide(AlignShapesRightDefaultNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignTopDefault()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.PowerpointDefaults;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            PositionsLabMain.AlignTop(actualShapes);

            PpOperations.SelectSlide(AlignShapesTopDefaultNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignBottomDefault()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.PowerpointDefaults;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideHeight = Pres.PageSetup.SlideHeight;
            PositionsLabMain.AlignBottom(actualShapes, slideHeight);

            PpOperations.SelectSlide(AlignShapesBottomDefaultNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignHorizontalDefault()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.PowerpointDefaults;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideHeight = Pres.PageSetup.SlideHeight;
            PositionsLabMain.AlignHorizontalCenter(actualShapes, slideHeight);

            PpOperations.SelectSlide(AlignShapesHorizontalDefaultNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignVerticalDefault()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.PowerpointDefaults;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideWidth = Pres.PageSetup.SlideWidth;
            PositionsLabMain.AlignVerticalCenter(actualShapes, slideWidth);

            PpOperations.SelectSlide(AlignShapesVerticalDefaultNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignCenterDefault()
        {
            PositionsLabSettings.AlignReference = PositionsLabSettings.AlignReferenceObject.PowerpointDefaults;
            _shapeNames = new List<string> { UnrotatedRectangle, Oval, RotatedArrow, RotatedRectangle };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            float slideHeight = Pres.PageSetup.SlideHeight;
            float slideWidth = Pres.PageSetup.SlideWidth;
            PositionsLabMain.AlignCenter(actualShapes, slideHeight, slideWidth);

            PpOperations.SelectSlide(AlignShapesCenterDefaultNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestAlignRadial()
        {
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.AlignRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(AlignShapesRadialNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
