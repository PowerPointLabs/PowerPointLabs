using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabSoftEdgeEffecTest : BaseSyncLabTest
    {
        private const int OriginalSlideNo = 1;
        private const int DesiredSlideNo = 2;
        
        private const string SourceShape = "Source";
        private const string DestinationShape = "Destination";
        
        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_SoftEdgeEffect.pptx";
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncSoftEdgeEffect()
        {
            SyncSoftEdgeEffect(DestinationShape, OriginalSlideNo, DesiredSlideNo);
        }

        private void SyncSoftEdgeEffect(string shapeToBeSynced, int sourceSlideNumber, int expectedSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(sourceSlideNumber, SourceShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(sourceSlideNumber, shapeToBeSynced);
            new SoftEdgeEffectFormat().SyncFormat(formatShape, newShape);

            CompareSlides(sourceSlideNumber, expectedSlideNo);
        }

    }
}
