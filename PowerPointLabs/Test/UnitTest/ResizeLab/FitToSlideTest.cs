using System.Collections.Generic;
using System.Windows;
using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ResizeLab;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class FitToSlideTest : BaseResizeLabTest
    {
        private readonly ResizeLabMain _resizeLab = new ResizeLabMain();
        private List<string> _shapeNames;

        private const string UnrotatedShapeName = "rectangle";
        private const string RotatedShapeName = "rotatedRectangle";

        private const int OriginalShapesSlideNo = 17;
        private const int FitToWidthSlideNo = 18;
        private const int FitToWidthAspectRatioSlideNo = 19;
        private const int FitToHeightSlideNo = 20;
        private const int FitToHeightAspectRatioSlideNo = 21;
        private const int FitToFillSlideNo = 22;
        private const int FitToFillAspectRatioSlideNo = 23;

        [TestInitialize]
        public void TestInitialize()
        {
            _shapeNames = new List<string> { UnrotatedShapeName, RotatedShapeName };
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFitToWidthWithoutAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;
            _resizeLab.FitToWidth(actualShapes, slideWidth, slideHeight, false);

            PpOperations.SelectSlide(FitToWidthSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFitToWidthWithAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;
            _resizeLab.FitToWidth(actualShapes, slideWidth, slideHeight, true);

            PpOperations.SelectSlide(FitToWidthAspectRatioSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFitToHeightWithoutAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;
            _resizeLab.FitToHeight(actualShapes, slideWidth, slideHeight, false);

            PpOperations.SelectSlide(FitToHeightSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFitToHeightWithAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;
            _resizeLab.FitToHeight(actualShapes, slideWidth, slideHeight, true);

            PpOperations.SelectSlide(FitToHeightAspectRatioSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFitToFillWithoutAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;
            _resizeLab.FitToFill(actualShapes, slideWidth, slideHeight, false);

            PpOperations.SelectSlide(FitToFillSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFitToFillWithAspectRatio()
        {
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;
            _resizeLab.FitToFill(actualShapes, slideWidth, slideHeight, true);

            PpOperations.SelectSlide(FitToFillAspectRatioSlideNo);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}