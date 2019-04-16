using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabFontTest : BaseSyncLabTest
    {
        private const int OriginalShapesSlideNo = 21;
        private const string CopyFromSmallShape = "CopyFromSmall";
        private const string CopyFromLargeShape = "CopyFromLarge";
        private const string CopyToShape = "Rectangle 3";

        //Results of Operations
        private const int SyncFontFamilySlideNo = 22;
        private const int SyncFontSizeSlideNo = 23;
        private const int SyncFontFillSlideNo = 24;
        private const int SyncOneFontStyleSlideNo = 25;
        private const int SyncAllFontStyleSlideNo = 26;

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncFontFamily()
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(OriginalShapesSlideNo, CopyFromLargeShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(OriginalShapesSlideNo, CopyToShape);
            new FontFormat().SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncFontFamilySlideNo);
            CheckFontStyle(OriginalShapesSlideNo, SyncFontFamilySlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncFontSize()
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(OriginalShapesSlideNo, CopyFromLargeShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(OriginalShapesSlideNo, CopyToShape);
            new FontSizeFormat().SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncFontSizeSlideNo);
            CheckFontStyle(OriginalShapesSlideNo, SyncFontSizeSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncFontFill()
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(OriginalShapesSlideNo, CopyFromLargeShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(OriginalShapesSlideNo, CopyToShape);
            new FontColorFormat().SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncFontFillSlideNo);
            CheckFontStyle(OriginalShapesSlideNo, SyncFontFillSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncOneFontStyle()
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(OriginalShapesSlideNo, CopyFromLargeShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(OriginalShapesSlideNo, CopyToShape);
            new FontStyleFormat().SyncFormat(formatShape, newShape);

            CheckFontStyle(OriginalShapesSlideNo, SyncOneFontStyleSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncAllFontStyle()
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(OriginalShapesSlideNo, CopyFromSmallShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(OriginalShapesSlideNo, CopyToShape);
            new FontStyleFormat().SyncFormat(formatShape, newShape);

            CheckFontStyle(OriginalShapesSlideNo, SyncAllFontStyleSlideNo);
        }

        //Changes in font style are too minute for CompareSlide to detect so we need to check them manually
        protected void CheckFontStyle(int actualShapesSlideNo, int expectedShapesSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Shape actualShape = GetShape(actualShapesSlideNo, CopyToShape);
            Microsoft.Office.Interop.PowerPoint.Shape expectedShape = GetShape(expectedShapesSlideNo, CopyToShape);

            Microsoft.Office.Interop.PowerPoint.Font actualFont = actualShape.TextFrame.TextRange.Font;
            Microsoft.Office.Interop.PowerPoint.Font expectedFont = expectedShape.TextFrame.TextRange.Font;

            Assert.IsTrue(actualFont.Bold == expectedFont.Bold
                && actualFont.Italic == expectedFont.Italic
                && actualFont.Underline == expectedFont.Underline,
                "Font Style does not match expected font style. Expected bold:{0}, italic: {1}, underline: {2}."
                    + "Actual bold:{3}, italic:{4}, underline: {5}",
                expectedFont.Bold, expectedFont.Italic, expectedFont.Underline,
                actualFont.Bold, actualFont.Italic, actualFont.Underline);
        }
    }
}
