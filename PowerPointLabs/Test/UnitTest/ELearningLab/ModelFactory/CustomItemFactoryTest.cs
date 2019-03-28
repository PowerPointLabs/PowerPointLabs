using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;

using PowerPointLabs.ELearningLab.ELearningWorkspace.ModelFactory;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;

namespace Test.UnitTest.ELearningLab.ModelFactory
{
   // AbstractItemFactory _customItemFactory;
    [TestClass]
    public class CustomItemFactoryTest
    {
        AbstractItemFactory _factory;
        List<CustomEffect> _effects;

        [TestInitialize]
        public void Init()
        {
            _effects = new List<CustomEffect>();
            _effects.Add(new CustomEffect("TestShape1", "TestShape1ID", AnimationType.Emphasis));
            _effects.Add(new CustomEffect("TestShape2", "TestShape2ID", AnimationType.Exit));
            _factory = new CustomItemFactory(_effects);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetCustomItemBlock()
        {
            CustomClickItem item = _factory.GetBlock() as CustomClickItem;
            IEnumerable<CustomSubItem> customSubItems = item.CustomItems;

            Assert.AreEqual(customSubItems.Count(), _effects.Count);

            Assert.AreEqual("TestShape1", customSubItems.ElementAt(0).ShapeName);
            Assert.AreEqual("TestShape1ID", customSubItems.ElementAt(0).ShapeId);
            Assert.AreEqual(AnimationType.Emphasis, customSubItems.ElementAt(0).Type);

            Assert.AreEqual("TestShape2", customSubItems.ElementAt(1).ShapeName);
            Assert.AreEqual("TestShape2ID", customSubItems.ElementAt(1).ShapeId);
            Assert.AreEqual(AnimationType.Exit, customSubItems.ElementAt(1).Type);
        }
    }
}
