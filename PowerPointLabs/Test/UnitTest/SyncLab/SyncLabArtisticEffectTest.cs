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
        
        private const string SourceShape = "Picture 2";
        private const string SetArtisticEffect = "Picture 9";
        private const string ReplacedArtisticEffect = "Picture 10";
        
        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_ArtisticEffect.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSetArtisticEffect()
        {
            SyncArtisticEffect(SetArtisticEffect, OriginalSetEffectSlideNo, DesiredSetEffectSlideNo);
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestReplaceArtisticEffect()
        {
            SyncArtisticEffect(ReplacedArtisticEffect, OriginalReplaceEffectSlideNo, DesiredReplaceEffectSlideNo);
        }

        private void SyncArtisticEffect(string shapeToSync, int sourceSlideNumber, int expectedSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(sourceSlideNumber, SourceShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(sourceSlideNumber, shapeToSync);
            new ArtisticEffectFormat().SyncFormat(formatShape, newShape);

            CompareSlides(sourceSlideNumber, expectedSlideNo);
        }

    }
}
