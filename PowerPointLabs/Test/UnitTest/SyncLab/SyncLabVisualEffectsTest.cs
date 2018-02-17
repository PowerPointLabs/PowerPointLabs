using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    /// <summary>
    /// DO NOT use PPT2013 to edit the source powerpoint.
    /// PPT2013 exhibits strange behavior on ArtisticEffect.
    /// ArtisticEffect is sometimes made permenant on pictures after saving the file.
    /// i.e, image's artistic effect cannot be removed/undone
    /// </summary>
    [TestClass]
    public class SyncLabVisualEffectsTest : BaseSyncLabTest
    {
        private const int OriginalReplaceEffectSlideNo = 4;
        private const int OriginalSetEffectSlideNo = 6;
        private const int DesiredReplaceEffectSlideNo = 7;
        private const int DesiredSetEffectSlideNo = 5;
        
        private const string SourceShape = "Picture 2";
        private const string SetArtisticEffect = "Picture 9";
        private const string ReplacedArtisticEffect = "Picture 10";
        
        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab_VisualEffects.pptx";
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestReplaceArtisticEffect()
        {
            SyncArtisticEffect(ReplacedArtisticEffect, OriginalReplaceEffectSlideNo, DesiredReplaceEffectSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSetArtisticEffect()
        {
            SyncArtisticEffect(SetArtisticEffect, OriginalSetEffectSlideNo, DesiredSetEffectSlideNo);
        }

        private void SyncArtisticEffect(string shapeToPaste, int sourceSlideNumber, int expectedSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(sourceSlideNumber, SourceShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(sourceSlideNumber, shapeToPaste);
            PictureEffectsFormat.SyncFormat(formatShape, newShape);

            CompareSlides(sourceSlideNumber, expectedSlideNo);
        }

    }
}
