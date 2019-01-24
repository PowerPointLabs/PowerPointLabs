using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class ReflectionEffectTest : BaseSyncLabTest
    {
        private const int OriginalDistanceSlideNo = 1;
        private const int DesiredDistanceSlideNo = 2;
        private const int OriginalBlurSlideNo = 3;
        private const int DesiredBlurSlideNo = 4;
        private const int OriginalSizeSlideNo = 5;
        private const int DesiredSizeSlideNo = 6;
        private const int OriginalTransparencySlideNo = 7;
        private const int DesiredTransparencySlideNo = 8;
        
        private const string SourceShape = "Source";
        private const string DestinationShape = "Destination";
        
        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_ReflectionEffect.pptx";
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncReflectionDistance()
        {
            SyncAndCompareFormat(DestinationShape, OriginalDistanceSlideNo, DesiredDistanceSlideNo, new ReflectionDistanceFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncReflectionBlur()
        {
            SyncAndCompareFormat(DestinationShape, OriginalBlurSlideNo, DesiredBlurSlideNo, new ReflectionBlurFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncReflectionSize()
        {
            SyncAndCompareFormat(DestinationShape, OriginalSizeSlideNo, DesiredSizeSlideNo, new ReflectionSizeFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncReflectionTransparency()
        {
            SyncAndCompareFormat(DestinationShape, OriginalTransparencySlideNo, DesiredTransparencySlideNo, new ReflectionTransparencyFormat());
        }

        private void SyncAndCompareFormat(string shapeToBeSynced, int sourceSlideNumber, int expectedSlideNo, Format format)
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(sourceSlideNumber, SourceShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(sourceSlideNumber, shapeToBeSynced);
            format.SyncFormat(formatShape, newShape);

            CompareSlides(sourceSlideNumber, expectedSlideNo, 0.99);
        }

    }
}
