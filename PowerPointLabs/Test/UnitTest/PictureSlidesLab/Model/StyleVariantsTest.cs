using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PictureSlidesLab.Model;

namespace Test.UnitTest.PictureSlidesLab.Model
{
    [TestClass]
    public class StyleVariantsTest
    {
        private StyleVariant variant;

        [TestInitialize]
        public void Init()
        {
            variant = new StyleVariant(new Dictionary<string, object>());
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestApply()
        {
            variant.Set("OptionName", "test option name");
            StyleOption option = new StyleOption();
            variant.Apply(option);
            Assert.AreEqual("test option name", option.OptionName);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestCopy()
        {
            variant.Set("TextBoxPosition", 999);
            variant.Set("OptionName", "test option name");
            variant = variant.Copy(new StyleOption());

            StyleOption option = new StyleOption();
            option.TextBoxPosition = 999;
            option.OptionName = "test option name";

            variant.Apply(option);
            Assert.AreEqual(5, option.TextBoxPosition);
            Assert.AreEqual("Reloaded", option.OptionName);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestIsNoEffect()
        {
            variant.Set("TextBoxPosition", 999);
            variant.Set("OptionName", "test option name");

            StyleOption option = new StyleOption();
            option.TextBoxPosition = 999;
            option.OptionName = "test option name";

            Assert.IsTrue(variant.IsNoEffect(option));

            option.TextBoxPosition = 4;

            Assert.IsFalse(variant.IsNoEffect(option));
        }
    }
}
