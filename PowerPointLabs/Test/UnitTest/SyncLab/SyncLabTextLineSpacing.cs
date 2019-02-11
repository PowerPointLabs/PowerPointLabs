using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncTestLineSpacing : BaseSyncLabTest
    {
        private const int SpacingSlidesNo = 12;
        private const int SpacingExpectedSlidesNo = 13;
        private const string SourceSuffix = " source";
        private const string TargetSuffix = " target";

        //Types of text formats
        private const string Single = "Single";
        private const string S1_5 = "1_5";
        private const string Double = "Double";
        private const string Exactly40 = "Exactly40";
        private const string Multiple3 = "Multiple3";

        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_Text.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncLineSpacing()
        {
            string[] prefixes = { Single, S1_5, Double, Exactly40, Multiple3 };
            foreach (string prefix in prefixes)
            {
                Microsoft.Office.Interop.PowerPoint.Shape formatShape =
                    GetShape(SpacingSlidesNo, prefix + SourceSuffix);

                Microsoft.Office.Interop.PowerPoint.Shape newShape =
                    GetShape(SpacingSlidesNo, prefix + TargetSuffix);
                new TextLineSpacingFormat().SyncFormat(formatShape, newShape);
                CheckTextLineSpacing(prefix + TargetSuffix,
                    SpacingSlidesNo, SpacingExpectedSlidesNo);
            }

            CompareSlides(SpacingSlidesNo, SpacingExpectedSlidesNo);
        }

        //Changes in text alignment are too minute for CompareSlide to detect so we need to check them manually
        protected void CheckTextLineSpacing(string shape, int actualShapesSlideNo, int expectedShapesSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Shape actualShape = GetShape(actualShapesSlideNo, shape);
            Microsoft.Office.Interop.PowerPoint.Shape expectedShape = GetShape(expectedShapesSlideNo, shape);

            Microsoft.Office.Core.ParagraphFormat2 actualPFormat = actualShape.TextFrame2.TextRange.ParagraphFormat;
            Microsoft.Office.Core.ParagraphFormat2 expectedPFormat = expectedShape.TextFrame2.TextRange.ParagraphFormat;

            Assert.IsTrue(actualPFormat.SpaceBefore == expectedPFormat.SpaceBefore &&
                actualPFormat.SpaceWithin == expectedPFormat.SpaceWithin &&
                actualPFormat.SpaceAfter == expectedPFormat.SpaceAfter,
                "Text line spacing does not match expected line spacing." +
                "Expected before:{0}, within:{1}, after:{2}. " +
                "Actual before:{3}, within:{4}, after:{5}.",
                expectedPFormat.SpaceBefore, expectedPFormat.SpaceWithin, expectedPFormat.SpaceAfter,
                actualPFormat.SpaceBefore, actualPFormat.SpaceWithin, actualPFormat.SpaceAfter);
        }
    }
}
