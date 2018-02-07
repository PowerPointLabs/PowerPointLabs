using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    /// <summary>
    /// DO NOT use PPT2013 to edit the source powerpoint.
    /// PPT2013 exhibits strange behavior on ArtisticEffect.
    /// ArtisticEffect is sometimes made permenant on pictures after saving the file.
    /// </summary>
    [TestClass]
    public class SyncLabVisualEffectsTest : BaseSyncLabTest
    {
        private const int OriginalShapesSlideNo = 4;
        private const int ReplaceEffectSlideNo = 6;
        private const int SetEffectSlideNo = 5;
        
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
            SyncArtisticEffect(ReplacedArtisticEffect, ReplaceEffectSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSetArtisticEffect()
        {
            SyncArtisticEffect(SetArtisticEffect, SetEffectSlideNo);
        }

        private void SyncArtisticEffect(string shapeToPaste, int expectedSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Shape formatShape = GetShape(OriginalShapesSlideNo, SourceShape);

            Microsoft.Office.Interop.PowerPoint.Shape newShape = GetShape(OriginalShapesSlideNo, shapeToPaste);
            PictureEffectsFormat.SyncFormat(formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, expectedSlideNo);
        }

    }
}