using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabGlowEffectTest: BaseSyncLabTest
    {
        private const int OriginalShadowSlideNo = 1;
        private const int DesiredShadowSlideNo = 2;
        
        private const string SourceShape = "Source";
        private const string DestinationShape = "Destination";
        
        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_GlowEffect.pptx";
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncGlowEffect()
        {
            SyncGlowEffect(DestinationShape, OriginalShadowSlideNo, DesiredShadowSlideNo);
        }

        private void SyncGlowEffect(string shapeToBeSynced, int sourceSlideNumber, int expectedSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(sourceSlideNumber, SourceShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(sourceSlideNumber, shapeToBeSynced);
            new GlowEffectFormat().SyncFormat(formatShape, newShape);

            CompareSlides(sourceSlideNumber, expectedSlideNo);
        }

    }
}