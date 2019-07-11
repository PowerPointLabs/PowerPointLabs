using System.Collections.Generic;
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
            PpOperations.MaximizeWindow();
            PpOperations.SelectSlide(TestSyncExplanationItemSlideNo);
            IELearningLabController eLearningLab = PplFeatures.ELearningLab;
            eLearningLab.OpenPane();
            ThreadUtil.WaitFor(5000);
            eLearningLab.AddSelfExplanationItem();
            TestSyncExplanationItems(eLearningLab);

            PpOperations.SelectSlide(TestReorderExplanationItemSlideNo);
            ThreadUtil.WaitFor(1000);
            eLearningLab.AddSelfExplanationItem();
            TestReorderExplanationItems(eLearningLab);

            PpOperations.SelectSlide(TestDeleteExplanationItemSlideNo);
            ThreadUtil.WaitFor(1000);
            eLearningLab.AddSelfExplanationItem();
            TestDeleteExplanationItems(eLearningLab);
            ThreadUtil.WaitFor(10000);
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_CreateExplanationBeforeTemplateTest()
        {
            IELearningLabController eLearningLab = PplFeatures.ELearningLab;
            eLearningLab.OpenPane();
            ThreadUtil.WaitFor(5000);
            ExplanationItemTemplate[] items = CreateStartItems();
            eLearningLab.CreateTemplateExplanations(items);
            eLearningLab.AddAbove(1);
            List<ExplanationItemTemplate> explanationItemTemplates = new List<ExplanationItemTemplate>(items);
            explanationItemTemplates.Insert(1, items[0]);
            AssertEqual(explanationItemTemplates.ToArray(), eLearningLab.GetExplanations());
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_CreateExplanationAfterTemplateTest()
        {
            IELearningLabController eLearningLab = PplFeatures.ELearningLab;
            eLearningLab.OpenPane();
            ThreadUtil.WaitFor(5000);
            ExplanationItemTemplate[] items = CreateStartItems();
            eLearningLab.CreateTemplateExplanations(items);
            eLearningLab.AddBelow(1);
            List<ExplanationItemTemplate> explanationItemTemplates = new List<ExplanationItemTemplate>(items);
            explanationItemTemplates.Insert(2, items[1]);
            AssertEqual(explanationItemTemplates.ToArray(), eLearningLab.GetExplanations());
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_CreateExplanationAtBottomTemplateTest()
        {
            IELearningLabController eLearningLab = PplFeatures.ELearningLab;
            eLearningLab.OpenPane();
            ThreadUtil.WaitFor(5000);
            ExplanationItemTemplate[] items = CreateStartItems();
            eLearningLab.CreateTemplateExplanations(items);
            eLearningLab.AddAtBottom();
            List<ExplanationItemTemplate> explanationItemTemplates = new List<ExplanationItemTemplate>(items);
            explanationItemTemplates.Add(items[items.Length - 1]);
            AssertEqual(explanationItemTemplates.ToArray(), eLearningLab.GetExplanations());
        }


        private void AssertEqual(object[] arr1, object[] arr2)
        {
            Assert.AreEqual(arr1.Length, arr2.Length);
            for (int i = 0; i < arr1.Length; i++)
            {
                Assert.AreEqual(arr1[i], arr2[i]);
            }
        }

        private ExplanationItemTemplate[] CreateStartItems()
        {
            ExplanationItemTemplate item1 = new ExplanationItemTemplate()
            {
                IsCallout = false,
                IsCaption = false,
                IsVoice = true,
                VoiceLabel = "",
                HasShortVersion = false,
                CaptionText = ""
            };
            ExplanationItemTemplate item2 = new ExplanationItemTemplate()
            {
                IsCallout = false,
                IsCaption = true,
                IsVoice = false,
                VoiceLabel = PplFeatures.ELearningLab.DefaultVoiceLabel,
                HasShortVersion = false,
                CaptionText = ""
            };
            ExplanationItemTemplate item3 = new ExplanationItemTemplate()
            {
                IsCallout = true,
                IsCaption = false,
                IsVoice = false,
                VoiceLabel = "",
                HasShortVersion = true,
                CaptionText = "Caption"
            };
            return new ExplanationItemTemplate[3] { item1, item2, item3 };
        }

        private void TestSyncExplanationItems(IELearningLabController eLearningLab)
        {
            eLearningLab.Sync();
            Slide expSlide = PpOperations.SelectSlide(TestSyncExplanationItemSlideNo);
            ThreadUtil.WaitFor(1000);
            Slide actualSlide = PpOperations.SelectSlide(ExpectedSyncExplanationItemSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide, similarityTolerance: 0.9);
        }

        private void TestReorderExplanationItems(IELearningLabController eLearningLab)
        {
            eLearningLab.Reorder();
            eLearningLab.Sync();
            Slide expSlide = PpOperations.SelectSlide(TestReorderExplanationItemSlideNo);
            ThreadUtil.WaitFor(1000);
            Slide actualSlide = PpOperations.SelectSlide(ExpectedReorderExplanationItemSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide, similarityTolerance: 0.9);
        }

        private void TestDeleteExplanationItems(IELearningLabController eLearningLab)
        {
            eLearningLab.Delete();
            eLearningLab.Sync();
            Slide expSlide = PpOperations.SelectSlide(TestDeleteExplanationItemSlideNo);
            ThreadUtil.WaitFor(1000);
            Slide actualSlide = PpOperations.SelectSlide(ExpectedDeleteExplanationItemSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide, similarityTolerance: 0.9);
        }
    }
}
