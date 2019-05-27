using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

using TestInterface;

namespace Test.FunctionalTest
{
    [TestClass]
    public class ELearningLabTest : BaseFunctionalTest
    {
        private const int TestSyncExplanationItemSlideNo = 1;
        private const int ExpectedSyncExplanationItemSlideNo = 2;
        private const int TestReorderExplanationItemSlideNo = 3;
        private const int ExpectedReorderExplanationItemSlideNo = 4;
        private const int TestDeleteExplanationItemSlideNo = 5;
        private const int ExpectedDeleteExplanationItemSlideNo = 6;

        protected override string GetTestingSlideName()
        {
            return "ELearningLab\\ELearningLabTest.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_CreateSelfExplanationTest()
        {
            PpOperations.SelectSlide(TestSyncExplanationItemSlideNo);
            IELearningLabController eLearningLab = PplFeatures.ELearningLab;
            eLearningLab.OpenPane();
            ThreadUtil.WaitFor(5000);
            eLearningLab.AddSelfExplanationItem();
            TestSyncExplanationItems(eLearningLab);

            PpOperations.SelectSlide(TestReorderExplanationItemSlideNo);
            ThreadUtil.WaitFor(5000);
            eLearningLab.AddSelfExplanationItem();
            TestReorderExplanationItems(eLearningLab);

            PpOperations.SelectSlide(TestDeleteExplanationItemSlideNo);
            ThreadUtil.WaitFor(5000);
            eLearningLab.AddSelfExplanationItem();
            TestDeleteExplanationItems(eLearningLab);
        }

        private void TestSyncExplanationItems(IELearningLabController eLearningLab)
        {
            eLearningLab.Sync();
            Slide expSlide = PpOperations.SelectSlide(TestSyncExplanationItemSlideNo);
            Slide actualSlide = PpOperations.SelectSlide(ExpectedSyncExplanationItemSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide, similarityTolerance: 0.9);
        }

        private void TestReorderExplanationItems(IELearningLabController eLearningLab)
        {
            eLearningLab.Reorder();
            eLearningLab.Sync();
            Slide expSlide = PpOperations.SelectSlide(TestReorderExplanationItemSlideNo);
            Slide actualSlide = PpOperations.SelectSlide(ExpectedReorderExplanationItemSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide, similarityTolerance: 0.9);
        }

        private void TestDeleteExplanationItems(IELearningLabController eLearningLab)
        {
            eLearningLab.Delete();
            eLearningLab.Sync();
            Slide expSlide = PpOperations.SelectSlide(TestDeleteExplanationItemSlideNo);
            Slide actualSlide = PpOperations.SelectSlide(ExpectedDeleteExplanationItemSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide, similarityTolerance: 0.9);
        }
    }
}
