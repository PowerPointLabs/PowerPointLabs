using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabGlowEffectTest: BaseSyncLabTest
    {
        private const int OriginalSizeSlideNo = 1;
        private const int DesiredSizeSlideNo = 2;
        private const int OriginalColorSlideNo = 3;
        private const int DesiredColorSlideNo = 4;
        private const int OriginalTransparencySlideNo = 5;
        private const int DesiredTransparencySlideNo = 6;
        
        private const string SourceShape = "Source";
        private const string DestinationShape = "Destination";
        
        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_GlowEffect.pptx";
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncGlowColor()
        {
            SyncAndCompareFormat(DestinationShape, OriginalColorSlideNo, DesiredColorSlideNo, new GlowColorFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncGlowSize()
        {
            SyncAndCompareFormat(DestinationShape, OriginalSizeSlideNo, DesiredSizeSlideNo, new GlowSizeFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncGlowTransparency()
        {
            SyncAndCompareFormat(DestinationShape, OriginalTransparencySlideNo, DesiredTransparencySlideNo, new GlowTransparencyFormat());
        }

        private void SyncAndCompareFormat(string shapeToBeSynced, int sourceSlideNumber, int expectedSlideNo, Format format)
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(sourceSlideNumber, SourceShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(sourceSlideNumber, shapeToBeSynced);
            format.SyncFormat(formatShape, newShape);

            CompareSlides(sourceSlideNumber, expectedSlideNo);
        }

    }
}