using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabThreeDRotationEffectTest : BaseSyncLabTest
    {
        private const int OriginalSlideNo = 1;
        private const int ExpectedSlideNo = 2;
        
        private const string SourceShape = "Source";
        private const string DestinationShape = "Destination";
        
        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_ThreeDRotationEffect.pptx";
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncThreeDRotation()
        {
            SyncThreeDRotation(DestinationShape, OriginalSlideNo, ExpectedSlideNo);
        }

        private void SyncThreeDRotation(string shapeToBeSynced, int sourceSlideNumber, int expectedSlideNo)
        {
            Shape formatShape = GetShape(sourceSlideNumber, SourceShape);

            Shape newShape = GetShape(sourceSlideNumber, shapeToBeSynced);
            new ThreeDRotationEffectFormat().SyncFormat(formatShape, newShape);

            CompareSlides(sourceSlideNumber, expectedSlideNo);
        }
    }
}
