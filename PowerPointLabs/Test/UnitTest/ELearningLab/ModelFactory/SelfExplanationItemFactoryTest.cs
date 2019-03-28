using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using PowerPointLabs.ELearningLab.ELearningWorkspace.ModelFactory;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;

namespace Test.UnitTest.ELearningLab.ModelFactory
{
    [TestClass]
    public class SelfExplanationItemFactoryTest
    { 
        [TestMethod]
        [TestCategory("UT")]
        public void TestGetSelfExplanationItemBlock()
        {
            List<ELLEffect> _effects = new List<ELLEffect>
            {
                new ELLEffect("PPTL_1_Callout"),
                new ELLEffect("PPTL_1_Caption"),
                new ELLEffect("PPTL_1_Audio_MichaelVoice")
            };

            AbstractItemFactory _factory = new SelfExplanationItemFactory(_effects);
            SelfExplanationClickItem item = _factory.GetBlock() as SelfExplanationClickItem;

            Assert.IsTrue(item.IsCallout);
            Assert.IsTrue(item.IsCaption);
            Assert.IsTrue(item.IsVoice);
            Assert.AreEqual(item.TagNo, 1);
            Assert.AreEqual(item.VoiceLabel, "MichaelVoice");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetSelfExplanationItem_NoCallout()
        {
            List<ELLEffect> _effects = new List<ELLEffect>
            {
                new ELLEffect("PPTL_1_Caption"),
                new ELLEffect("PPTL_1_Audio_MichaelVoice")
            };

            AbstractItemFactory _factory = new SelfExplanationItemFactory(_effects);
            SelfExplanationClickItem item = _factory.GetBlock() as SelfExplanationClickItem;

            Assert.IsFalse(item.IsCallout);
            Assert.IsTrue(item.IsCaption);
            Assert.IsTrue(item.IsVoice);
            Assert.AreEqual(item.TagNo, 1);
            Assert.AreEqual(item.VoiceLabel, "MichaelVoice");
        }
    }
}
