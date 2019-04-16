using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabTextAlignmentTest : BaseSyncLabTest
    {
        private const int HorizontalAlignmentSlidesNo = 4;
        private const int HorizontalExpectedSlidesNo = 5;
        private const int VerticalAlignmentSlidesNo = 6;
        private const int VerticalExpectedSlidesNo = 7;
        private const string SourceSuffix = " source";
        private const string TargetSuffix = " target";

        //Types of text formats
        private const string Left = "Left";
        private const string Center = "Center";
        private const string Right = "Right";
        private const string Justify = "Justify";
        private const string Top = "Top";
        private const string Middle = "Middle";
        private const string Bottom = "Bottom";

        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_Text.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncHorizontalAlignment()
        {
            string[] prefixes = { Left, Center, Right, Justify };
            foreach (string prefix in prefixes)
            {
                Microsoft.Office.Interop.PowerPoint.Shape formatShape =
                    GetShape(HorizontalAlignmentSlidesNo, prefix + SourceSuffix);

                Microsoft.Office.Interop.PowerPoint.Shape newShape =
                    GetShape(HorizontalAlignmentSlidesNo, prefix + TargetSuffix);
                new TextHorizontalAlignmentFormat().SyncFormat(formatShape, newShape);
                CheckTextAlignment(prefix + TargetSuffix,
                    HorizontalAlignmentSlidesNo, HorizontalExpectedSlidesNo);
            }

            CompareSlides(HorizontalAlignmentSlidesNo, HorizontalExpectedSlidesNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncVerticalAlignment()
        {
            string[] prefixes = { Top, Middle, Bottom };
            foreach (string prefix in prefixes)
            {
                Microsoft.Office.Interop.PowerPoint.Shape formatShape =
                    GetShape(VerticalAlignmentSlidesNo, prefix + SourceSuffix);

                Microsoft.Office.Interop.PowerPoint.Shape newShape =
                    GetShape(VerticalAlignmentSlidesNo, prefix + TargetSuffix);
                new TextVerticalAlignmentFormat().SyncFormat(formatShape, newShape);

                CheckTextAlignment(prefix + TargetSuffix,
                    VerticalAlignmentSlidesNo, VerticalExpectedSlidesNo);
            }

            CompareSlides(VerticalAlignmentSlidesNo, VerticalExpectedSlidesNo);
        }

        //Changes in text alignment are too minute for CompareSlide to detect so we need to check them manually
        protected void CheckTextAlignment(string shape, int actualShapesSlideNo, int expectedShapesSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Shape actualShape = GetShape(actualShapesSlideNo, shape);
            Microsoft.Office.Interop.PowerPoint.Shape expectedShape = GetShape(expectedShapesSlideNo, shape);

            Microsoft.Office.Interop.PowerPoint.TextFrame2 actualTextFrame = actualShape.TextFrame2;
            Microsoft.Office.Interop.PowerPoint.TextFrame2 expectedTextFrame = expectedShape.TextFrame2;

            Assert.IsTrue(actualTextFrame.TextRange.ParagraphFormat.Alignment == expectedTextFrame.TextRange.ParagraphFormat.Alignment
                && actualTextFrame.HorizontalAnchor == expectedTextFrame.HorizontalAnchor
                && actualTextFrame.VerticalAnchor == expectedTextFrame.VerticalAnchor,
                "Text alignment does not match expected text alignment." +
                "Expected paragraphAlignment: {0}, horizontalAlignment:{1}, verticalAlignment: {2}."
                    + "Actual paragraphAlignment: {3}, horizontalAlignment:{4}, verticalAlignment:{5}.",
                expectedTextFrame.TextRange.ParagraphFormat.Alignment,
                expectedTextFrame.HorizontalAnchor, expectedTextFrame.VerticalAnchor,
                actualTextFrame.TextRange.ParagraphFormat.Alignment,
                actualTextFrame.HorizontalAnchor, actualTextFrame.VerticalAnchor);
        }
    }
}
