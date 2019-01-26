using System;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabLineTest : BaseSyncLabTest
    {
        private const int OriginalShapesSlideNo = 28;
        private const string CopyFromShape = "CopyFrom";
        private const string StraightLine = "Straight Connector 2";
        private const string Arrow = "Right Arrow 5";

        //Results of Operations
        private const int SyncLineFillSlideNo = 29;
        private const int SyncLineWidthSlideNo = 30;
        private const int SyncLineCompoundTypeSlideNo = 31;
        private const int SyncLineDashTypeSlideNo = 32;
        private const int SyncLineArrowSlideNo = 33;
        private const int SyncLineTransparencySlideNo = 34;

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncLineFill()
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(OriginalShapesSlideNo, CopyFromShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(OriginalShapesSlideNo, StraightLine);
            new LineFillFormat().SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncLineFillSlideNo);
            CheckLineStyle(StraightLine, OriginalShapesSlideNo, SyncLineFillSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncLineWidth()
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(OriginalShapesSlideNo, CopyFromShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(OriginalShapesSlideNo, Arrow);
            new LineWeightFormat().SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncLineWidthSlideNo);
            CheckLineStyle(Arrow, OriginalShapesSlideNo, SyncLineWidthSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncLineCompoundType()
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(OriginalShapesSlideNo, CopyFromShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(OriginalShapesSlideNo, StraightLine);
            new LineCompoundTypeFormat().SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncLineCompoundTypeSlideNo);
            CheckLineStyle(StraightLine, OriginalShapesSlideNo, SyncLineCompoundTypeSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncLineDashType()
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(OriginalShapesSlideNo, CopyFromShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(OriginalShapesSlideNo, StraightLine);
            new LineDashTypeFormat().SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncLineDashTypeSlideNo);
            CheckLineStyle(StraightLine, OriginalShapesSlideNo, SyncLineDashTypeSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncLineArrow()
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(OriginalShapesSlideNo, CopyFromShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(OriginalShapesSlideNo, StraightLine);
            new LineArrowFormat().SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncLineArrowSlideNo);
            CheckLineStyle(StraightLine, OriginalShapesSlideNo, SyncLineArrowSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncLineTransparency()
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(OriginalShapesSlideNo, CopyFromShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(OriginalShapesSlideNo, Arrow);
            new LineTransparencyFormat().SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncLineTransparencySlideNo);
            CheckLineStyle(Arrow, OriginalShapesSlideNo, SyncLineTransparencySlideNo);
        }

        //Changes in line style are too minute for CompareSlide to detect so we need to check them manually
        protected void CheckLineStyle(string shape, int actualShapesSlideNo, int expectedShapesSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Shape actualShape = GetShape(actualShapesSlideNo, shape);
            Microsoft.Office.Interop.PowerPoint.Shape expectedShape = GetShape(expectedShapesSlideNo, shape);

            //Check transparency style
            Assert.IsTrue(Math.Abs(actualShape.Line.Transparency - expectedShape.Line.Transparency) < 0.001,
                "different transparency. exp:{0}, actual:{1}",
                expectedShape.Line.Transparency, actualShape.Line.Transparency);

            //Check line style
            Assert.IsTrue(actualShape.Line.Style == expectedShape.Line.Style,
                "different compound type. exp:{0}, actual:{1}",
                expectedShape.Line.Style, actualShape.Line.Style);
            Assert.IsTrue(actualShape.Line.DashStyle == expectedShape.Line.DashStyle,
                "different dash type. exp:{0}, actual:{1}",
                expectedShape.Line.DashStyle, actualShape.Line.DashStyle);

            //Check arrow style
            Assert.IsTrue(actualShape.Line.BeginArrowheadLength == expectedShape.Line.BeginArrowheadLength,
                "different begin arrowhead length. exp:{0}, actual:{1}",
                expectedShape.Line.BeginArrowheadLength, actualShape.Line.BeginArrowheadLength);
            Assert.IsTrue(actualShape.Line.BeginArrowheadStyle == expectedShape.Line.BeginArrowheadStyle,
                "different begin arrowhead style. exp:{0}, actual:{1}",
                expectedShape.Line.BeginArrowheadStyle, actualShape.Line.BeginArrowheadStyle);
            Assert.IsTrue(actualShape.Line.BeginArrowheadWidth == expectedShape.Line.BeginArrowheadWidth,
                "different begin arrowhead width. exp:{0}, actual:{1}",
                expectedShape.Line.BeginArrowheadWidth, actualShape.Line.BeginArrowheadWidth);
            Assert.IsTrue(actualShape.Line.EndArrowheadLength == expectedShape.Line.EndArrowheadLength,
                "different end arrowhead length. exp:{0}, actual:{1}",
                expectedShape.Line.EndArrowheadLength, actualShape.Line.EndArrowheadLength);
            Assert.IsTrue(actualShape.Line.EndArrowheadStyle == expectedShape.Line.EndArrowheadStyle,
                "different end arrowhead style. exp:{0}, actual:{1}",
                expectedShape.Line.EndArrowheadStyle, actualShape.Line.EndArrowheadStyle);
            Assert.IsTrue(actualShape.Line.EndArrowheadWidth == expectedShape.Line.EndArrowheadWidth,
                "different end arrowhead width. exp:{0}, actual:{1}",
                expectedShape.Line.EndArrowheadWidth, actualShape.Line.EndArrowheadWidth);
        }
    }
}
