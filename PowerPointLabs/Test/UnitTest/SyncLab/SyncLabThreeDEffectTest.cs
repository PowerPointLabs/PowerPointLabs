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

        private Format[] _threeDFormats =
        {
            new BevelBottomFormat(),
            new BevelTopFormat(),
            new ContourColorFormat(),
            new ContourWidthFormat(),
            new DepthColorFormat(), 
            new DepthSizeFormat(), 
            new LightingAngleFormat(), 
            new LightingEffectFormat(), 
            new MaterialEffectFormat() 
        };
        
        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_BevelEffect.pptx";
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncThreeDEffects()
        {
            SyncEffects(DestinationShape, OriginalSlideNo, ExpectedSlideNo);
        }

        private void SyncEffects(string shapeToBeSynced, int sourceSlideNumber, int expectedSlideNo)
        {
            Shape formatShape = GetShape(sourceSlideNumber, SourceShape);

            Shape newShape = GetShape(sourceSlideNumber, shapeToBeSynced);
            foreach (Format format in _threeDFormats)
            {
                format.SyncFormat(formatShape, newShape);
            }

            CompareSlides(sourceSlideNumber, expectedSlideNo);
        }

    }
}
