using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabBevelEffectTest : BaseSyncLabTest
    {
        private const int OriginalSlideNo = 1;
        private const int ExpectedSlideNo = 2;
        
        private const string SourceShape = "Source";
        private const string DestinationShape = "Destination";
        
        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_BevelEffect.pptx";
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncBevel()
        {
            SyncBevel(DestinationShape, OriginalSlideNo, ExpectedSlideNo);
        }

        private void SyncBevel(string shapeToBeSynced, int sourceSlideNumber, int expectedSlideNo)
        {
            Shape formatShape = GetShape(sourceSlideNumber, SourceShape);

            Shape newShape = GetShape(sourceSlideNumber, shapeToBeSynced);
            new BevelEffectFormat().SyncFormat(formatShape, newShape);

            CompareSlides(sourceSlideNumber, expectedSlideNo);
        }

    }
}
