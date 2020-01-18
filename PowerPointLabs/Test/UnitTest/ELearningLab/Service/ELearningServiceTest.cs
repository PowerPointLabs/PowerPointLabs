using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace Test.UnitTest.ELearningLab.Service
{
    [TestClass]
    public class ELearningServiceTest : BaseUnitTest
    {
        protected override string GetTestingSlideName()
        {
            return "ELearningLab\\ELearningServiceTest.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestRemoveLabAnimations()
        {
            PowerPointSlide slide = PowerPointSlide.FromSlideFactory(PpOperations.SelectSlide(1));
            ELearningService service = new ELearningService(slide, null);
            service.RemoveLabAnimationsFromAnimationPane();
            Assert.IsTrue(Util.SlideUtil.IsAnimationsRemoved(slide.GetNativeSlide(), ELearningLabText.Identifier));
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDeleteShapes()
        {

        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncAppearEffectsForExplanationItems()
        {

        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncExitEffectForExplanationItems()
        {

        }
    }
}
