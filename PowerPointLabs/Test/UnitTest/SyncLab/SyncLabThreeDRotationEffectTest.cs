using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabThreeDRotationEffectTest : BaseSyncLabTest
    {
        private const int OriginalSlideNo = 1;
        private const int ExpectedRotationXSlideNo = 2;
        private const int ExpectedRotationYSlideNo = 3;
        private const int ExpectedRotationZSlideNo = 4;
        private const int ExpectedPerspectiveSlideNo = 5;
        private const int ExpectedProjectTextSlideNo = 6;
        private const int ExpectedZSlideNo = 7;

        private const string SourceShape = "Source";
        private const string DestinationShape = "Destination";
        
        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_ThreeDRotationEffect.pptx";
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncThreeDRotationEffectRotationX()
        {
            SyncThreeDRotationEffect(DestinationShape, 
                OriginalSlideNo,
                ExpectedRotationXSlideNo, 
                new ThreeDRotationEffectRotationXFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncThreeDRotationEffectRotationY()
        {
            SyncThreeDRotationEffect(DestinationShape,
                OriginalSlideNo,
                ExpectedRotationYSlideNo,
                new ThreeDRotationEffectRotationYFormat());
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncThreeDRotationEffectRotationZ()
        {
            SyncThreeDRotationEffect(DestinationShape,
                OriginalSlideNo,
                ExpectedRotationZSlideNo,
                new ThreeDRotationEffectRotationZFormat());
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncThreeDRotationEffectPerspective()
        {
            SyncThreeDRotationEffect(DestinationShape,
                OriginalSlideNo,
                ExpectedPerspectiveSlideNo,
                new ThreeDRotationEffectPerspectiveFormat());
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncThreeDRotationEffectProjectText()
        {
            SyncThreeDRotationEffect(DestinationShape,
                OriginalSlideNo,
                ExpectedProjectTextSlideNo,
                new ThreeDRotationEffectProjectTextFormat());
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncThreeDRotationEffectZ()
        {
            SyncThreeDRotationEffect(DestinationShape,
                OriginalSlideNo,
                ExpectedZSlideNo,
                new ThreeDRotationEffectZFormat());
        }

        private void SyncThreeDRotationEffect(string shapeToBeSynced, int sourceSlideNumber, int expectedSlideNo, Format format)
        {
            Shape formatShape = GetShape(sourceSlideNumber, SourceShape);
            Shape newShape = GetShape(sourceSlideNumber, shapeToBeSynced);
            format.SyncFormat(formatShape, newShape);

            CompareSlides(sourceSlideNumber, expectedSlideNo);
        }
    }
}
