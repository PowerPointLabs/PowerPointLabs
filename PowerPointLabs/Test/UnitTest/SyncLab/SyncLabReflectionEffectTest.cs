using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.SyncLab.ObjectFormats;
using Microsoft.Office.Interop.PowerPoint;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class ReflectionEffectTest : BaseSyncLabTest
    {
        private const int OriginalSlideNo = 1;
        private const int ExpectedSlideNo = 2;
        
        private const string SourceShape = "Source";
        private const string DestinationShape = "Destination";
        
        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_ReflectionEffect.pptx";
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncReflection()
        {
            SyncReflection(DestinationShape, OriginalSlideNo, ExpectedSlideNo);
        }

        private void SyncReflection(string shapeToBeSynced, int sourceSlideNumber, int expectedSlideNo)
        {
            Shape formatShape = GetShape(sourceSlideNumber, SourceShape);

            Shape newShape = GetShape(sourceSlideNumber, shapeToBeSynced);
            new ReflectionEffectFormat().SyncFormat(formatShape, newShape);

            CompareSlides(sourceSlideNumber, expectedSlideNo);
        }

    }
}
