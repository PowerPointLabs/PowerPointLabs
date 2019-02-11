using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabTextOrientation : BaseSyncLabTest
    {
        private const int OrientationSlidesNo = 9;
        private const int OrientationExpectedSlidesNo = 10;
        private const string SourceSuffix = " source";
        private const string TargetSuffix = " target";

        //Types of text formats
        private const string Horizontal = "Horizontal";
        private const string Vertical = "Vertical";
        private const string Rotate90 = "Rotate90";
        private const string Rotate270 = "Rotate270";
        private const string Stacked = "Stacked";

        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_Text.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncOrientation()
        {
            string[] prefixes = { Horizontal, Vertical, Rotate90, Rotate270, Stacked };
            foreach (string prefix in prefixes)
            {
                Microsoft.Office.Interop.PowerPoint.Shape formatShape =
                    GetShape(OrientationSlidesNo, prefix + SourceSuffix);

                Microsoft.Office.Interop.PowerPoint.Shape newShape =
                    GetShape(OrientationSlidesNo, prefix + TargetSuffix);
                new TextOrientationFormat().SyncFormat(formatShape, newShape);
                CheckTextOrientation(prefix + TargetSuffix,
                    OrientationSlidesNo, OrientationExpectedSlidesNo);
            }

            CompareSlides(OrientationSlidesNo, OrientationExpectedSlidesNo);
        }

        //Changes in text alignment are too minute for CompareSlide to detect so we need to check them manually
        protected void CheckTextOrientation(string shape, int actualShapesSlideNo, int expectedShapesSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Shape actualShape = GetShape(actualShapesSlideNo, shape);
            Microsoft.Office.Interop.PowerPoint.Shape expectedShape = GetShape(expectedShapesSlideNo, shape);

            Microsoft.Office.Interop.PowerPoint.TextFrame2 actualTextFrame = actualShape.TextFrame2;
            Microsoft.Office.Interop.PowerPoint.TextFrame2 expectedTextFrame = expectedShape.TextFrame2;

            Assert.IsTrue(actualTextFrame.Orientation == expectedTextFrame.Orientation,
                "Text orientation does not match expected text orientation." +
                "Expected orientation:{0}. Actual orientation:{1}.",
                expectedTextFrame.Orientation, actualTextFrame.Orientation);
        }
    }
}
