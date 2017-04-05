using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
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
            syncFill(SolidFill, SyncSolidFillSlideNo);
            syncFill(PatternFill, SyncPatternFillSlideNo);
            syncFill(BackgroundFill, SyncBackgroundFillSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncGradientFill()
        {
            syncFill(TwoGradientFill, SyncTwoGradientFillSlideNo);
            syncFill(DiagonalGradientFill, SyncDiagonalGradientFillSlideNo);
            syncFill(RectangularGradientFill, SyncRectangularGradientFillSlideNo);
            syncFill(PathGradientFill, SyncPathGradientFillSlideNo);
            syncFill(NineGradientFill, SyncNineGradientFillSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncTransparency()
        {
            var formatShape = GetShape(OriginalShapesSlideNo, SolidFill);

            var newShape = GetShape(OriginalShapesSlideNo, CopyToShape);
            FillTransparencyFormat.SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncTransparencySlideNo);
            CheckTransparency(CopyToShape, OriginalShapesSlideNo, SyncTransparencySlideNo);
        }

        protected void syncFill(string shapeToCopy, int expectedSlideNo)
        {
            var formatShape = GetShape(OriginalShapesSlideNo, shapeToCopy);

            var newShape = GetShape(OriginalShapesSlideNo, CopyToShape);
            FillFormat.SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, expectedSlideNo);
            CheckTransparency(CopyToShape, OriginalShapesSlideNo, expectedSlideNo);
        }

        //Changes in transparency are too minute for CompareSlide to detect so we need to check them manually
        protected void CheckTransparency(string shape, int actualShapesSlideNo, int expectedShapesSlideNo)
        {
            var actualShape = GetShape(actualShapesSlideNo, shape);
            var expectedShape = GetShape(expectedShapesSlideNo, shape);

            Assert.IsTrue(actualShape.Fill.Transparency == expectedShape.Fill.Transparency,
                "different transparency. exp:{0}, actual:{1}",
                expectedShape.Fill.Transparency, actualShape.Fill.Transparency);
        }
    }
}
