using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    /// <summary>
    /// Be wary when using PPT2013 to edit the source powerpoint.
    /// ArtisticEffect is sometimes made permenant on PPT2013
    /// i.e, image's artistic effect cannot be removed/undone after the image is copied
    /// </summary>
    [TestClass]
    public class SyncLabArtisticEffectTest : BaseSyncLabTest
    {
        private const int OriginalSetEffectSlideNo = 4;
        private const int DesiredSetEffectSlideNo = 5;
        private const int OriginalReplaceEffectSlideNo = 6;
        private const int DesiredReplaceEffectSlideNo = 7;
        
        private const string SourceShape = "Source";
        private const string DestinationShape = "Destination";
        
        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_ArtisticEffect.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSetArtisticEffect()
        {
            SyncAndCompareFormat(DestinationShape, OriginalSetEffectSlideNo, DesiredSetEffectSlideNo, new ArtisticEffectFormat());
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestReplaceArtisticEffect()
        {
            SyncAndCompareFormat(DestinationShape, OriginalReplaceEffectSlideNo, DesiredReplaceEffectSlideNo, new ArtisticEffectFormat());
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
