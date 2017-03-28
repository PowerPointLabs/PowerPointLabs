using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabFillTest : BaseSyncLabTest
    {
        private const int OriginalShapesSlideNo = 10;
        private const string SolidFill = "SolidShape";
        private const string PatternFill = "PatternShape";
        private const string BackgroundFill = "BackgroundShape";
        private const string CopyToShape = "Rectangle 3";

        //Results of Operations
        private const int SyncPatternFillSlideNo = 11;
        private const int SyncSolidFillSlideNo = 12;
        private const int SyncBackgroundFillSlideNo = 13;

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncSolidFill()
        {
            syncFill(SolidFill, SyncSolidFillSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncPatternFill()
        {
            syncFill(PatternFill, SyncPatternFillSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncBackgroundFill()
        {
            syncFill(BackgroundFill, SyncBackgroundFillSlideNo);
        }

        protected void syncFill(string shapeToCopy, int expectedSlideNo)
        {
            var formatShape = GetShape(OriginalShapesSlideNo, shapeToCopy);

            var newShape = GetShape(OriginalShapesSlideNo, CopyToShape);
            FillFormat.SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, expectedSlideNo);
        }
    }
}
