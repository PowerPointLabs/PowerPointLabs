using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabBevelEffectTest : BaseSyncLabTest
    {
        private const int BevelTopOriginalSlideNo = 1;
        private const int BevelTopDesiredSlideNo = 2;
        private const int BevelBottomOriginalSlideNo = 3;
        private const int BevelBottomDesiredSlideNo = 4;
        private const int ContourColorOriginalSlideNo = 5;
        private const int ContourColorDesiredSlideNo = 6;
        private const int ContourWidthOriginalSlideNo = 7;
        private const int ContourWidthDesiredSlideNo = 8;
        private const int DepthColorOriginalSlideNo = 9;
        private const int DepthColorDesiredSlideNo = 10;
        private const int DepthSizeOriginalSlideNo = 11;
        private const int DepthSizeDesiredSlideNo = 12;
        private const int LightingEffectOriginalSlideNo = 13;
        private const int LightingEffectDesiredSlideNo = 14;
        private const int LightingAngleOriginalSlideNo = 15;
        private const int LightingAngleDesiredSlideNo = 16;
        private const int MaterialEffectOriginalSlideNo = 17;
        private const int MaterialEffectDesiredSlideNo = 18;
        
        private const string SourceShape = "Source";
        private const string DestinationShape = "Destination";

        
        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_ThreeDEffects.pptx";
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncBevelTop()
        {
            SyncEffects(BevelTopOriginalSlideNo,
                BevelTopDesiredSlideNo, 
                new BevelTopFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncBevelBottom()
        {
            SyncEffects(BevelBottomOriginalSlideNo,
                BevelBottomDesiredSlideNo, 
                new BevelBottomFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncContourColor()
        {
            SyncEffects(ContourColorOriginalSlideNo,
                ContourColorDesiredSlideNo, 
                new ContourColorFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncContourWidth()
        {
            SyncEffects(ContourWidthOriginalSlideNo,
                ContourWidthDesiredSlideNo, 
                new ContourWidthFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncDepthColor()
        {
            SyncEffects(DepthColorOriginalSlideNo,
                DepthColorDesiredSlideNo, 
                new DepthColorFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncDepthSize()
        {
            SyncEffects(DepthSizeOriginalSlideNo,
                DepthSizeDesiredSlideNo, 
                new DepthSizeFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncLightingEffect()
        {
            SyncEffects(LightingEffectOriginalSlideNo,
                LightingEffectDesiredSlideNo, 
                new LightingEffectFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncLightingAngle()
        {
            SyncEffects(LightingAngleOriginalSlideNo,
                LightingAngleDesiredSlideNo, 
                new LightingAngleFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncMaterialEffect()
        {
            SyncEffects(MaterialEffectOriginalSlideNo,
                MaterialEffectDesiredSlideNo, 
                new MaterialEffectFormat());
        }
        
        private void SyncEffects(int sourceSlideNumber, int expectedSlideNo, Format format)
        {
            Shape formatShape = GetShape(sourceSlideNumber, SourceShape);
            Shape newShape = GetShape(sourceSlideNumber, DestinationShape);
            format.SyncFormat(formatShape, newShape);

            CompareSlides(sourceSlideNumber, expectedSlideNo, 0.99);
        }

    }
}
