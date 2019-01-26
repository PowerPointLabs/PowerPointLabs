using System.Collections.Generic;

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

        [TestInitialize]
        public void TestInitialize()
        {
            _shapeNames = new List<string> { UnrotatedShapeName, RotatedShapeName };
            InitOriginalShapes(SlideNo.FitToSlideOrigin, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(SlideNo.FitToSlideOrigin, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFitToWidthWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.FitToSlideOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.FitToSlideWidth, _shapeNames);
            float slideWidth = Pres.PageSetup.SlideWidth;
            float slideHeight = Pres.PageSetup.SlideHeight;

            _resizeLab.FitToWidth(actualShapes, slideWidth, slideHeight, false);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFitToWidthWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.FitToSlideOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.FitToSlideWidthAspectRatio, _shapeNames);
            float slideWidth = Pres.PageSetup.SlideWidth;
            float slideHeight = Pres.PageSetup.SlideHeight;

            _resizeLab.FitToWidth(actualShapes, slideWidth, slideHeight, true);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFitToHeightWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.FitToSlideOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.FitToSlideHeight, _shapeNames);
            float slideWidth = Pres.PageSetup.SlideWidth;
            float slideHeight = Pres.PageSetup.SlideHeight;

            _resizeLab.FitToHeight(actualShapes, slideWidth, slideHeight, false);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFitToHeightWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.FitToSlideOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.FitToSlideHeightAspectRatio, _shapeNames);
            float slideWidth = Pres.PageSetup.SlideWidth;
            float slideHeight = Pres.PageSetup.SlideHeight;

            _resizeLab.FitToHeight(actualShapes, slideWidth, slideHeight, true);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFitToFillWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.FitToSlideOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.FitToSlideFill, _shapeNames);
            float slideWidth = Pres.PageSetup.SlideWidth;
            float slideHeight = Pres.PageSetup.SlideHeight;

            _resizeLab.FitToFill(actualShapes, slideWidth, slideHeight, false);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFitToFillWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.FitToSlideOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.FitToSlideFillAspectRatio, _shapeNames);
            float slideWidth = Pres.PageSetup.SlideWidth;
            float slideHeight = Pres.PageSetup.SlideHeight;

            _resizeLab.FitToFill(actualShapes, slideWidth, slideHeight, true);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}