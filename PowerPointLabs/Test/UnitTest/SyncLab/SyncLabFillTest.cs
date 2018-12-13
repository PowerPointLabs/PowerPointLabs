using System;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabFillTest : BaseSyncLabTest
    {
        private const int OriginalShapesSlideNo = 10;
        private const string CopyToShape = "Rectangle 3";
        private const string SolidFill = "SolidShape";
        private const string PatternFill = "PatternShape";
        private const string BackgroundFill = "BackgroundShape";
        //As of PPT2016, nine gradient stops is the maximum
        private const string NineGradientFill = "NineGradientShape";
        private const string TwoGradientFill = "TwoGradientShape";
        private const string DiagonalGradientFill = "DiagonalGradientShape";
        private const string RectangularGradientFill = "RectangularGradientShape";
        private const string PathGradientFill = "PathGradientShape";

        //Results of Operations
        private const int SyncPatternFillSlideNo = 11;
        private const int SyncSolidFillSlideNo = 12;
        private const int SyncBackgroundFillSlideNo = 13;
        private const int SyncTwoGradientFillSlideNo = 14;
        private const int SyncDiagonalGradientFillSlideNo = 15;
        private const int SyncRectangularGradientFillSlideNo = 16;
        private const int SyncPathGradientFillSlideNo = 17;
        private const int SyncNineGradientFillSlideNo = 18;
        private const int SyncTransparencySlideNo = 19;

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncFill()
        {
            SyncFill(SolidFill, SyncSolidFillSlideNo);
            SyncFill(PatternFill, SyncPatternFillSlideNo);
            SyncFill(BackgroundFill, SyncBackgroundFillSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncGradientFill()
        {
            SyncFill(TwoGradientFill, SyncTwoGradientFillSlideNo);
            SyncFill(DiagonalGradientFill, SyncDiagonalGradientFillSlideNo);
            SyncFill(RectangularGradientFill, SyncRectangularGradientFillSlideNo);
            SyncFill(PathGradientFill, SyncPathGradientFillSlideNo);
            SyncFill(NineGradientFill, SyncNineGradientFillSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncTransparency()
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(OriginalShapesSlideNo, SolidFill);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(OriginalShapesSlideNo, CopyToShape);
            new FillTransparencyFormat().SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncTransparencySlideNo);
            CheckTransparency(CopyToShape, OriginalShapesSlideNo, SyncTransparencySlideNo);
        }

        protected void SyncFill(string shapeToCopy, int expectedSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(OriginalShapesSlideNo, shapeToCopy);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(OriginalShapesSlideNo, CopyToShape);
            new FillFormat().SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, expectedSlideNo);
            CheckTransparency(CopyToShape, OriginalShapesSlideNo, expectedSlideNo);
        }

        //Changes in transparency are too minute for CompareSlide to detect so we need to check them manually
        protected void CheckTransparency(string shape, int actualShapesSlideNo, int expectedShapesSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Shape actualShape = GetShape(actualShapesSlideNo, shape);
            Microsoft.Office.Interop.PowerPoint.Shape expectedShape = GetShape(expectedShapesSlideNo, shape);

            Assert.IsTrue(Math.Abs(actualShape.Fill.Transparency - expectedShape.Fill.Transparency) < 0.001,
                "different transparency. exp:{0}, actual:{1}",
                expectedShape.Fill.Transparency, actualShape.Fill.Transparency);
        }
    }
}
