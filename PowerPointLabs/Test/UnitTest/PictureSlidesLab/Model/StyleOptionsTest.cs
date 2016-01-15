using System;
using System.ComponentModel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using Test.Util;

namespace Test.UnitTest.PictureSlidesLab.Model
{
    [TestClass]
    public class StyleOptionsTest
    {
        private StyleOptions option;

        [TestInitialize]
        public void Init()
        {
            option = new StyleOptions();
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSerialization()
        {
            option.OptionName = "Test Option Name";
            option.Save(PathUtil.GetDocTestPath() + "option.user");
            var loadedOption = StyleOptions.Load(PathUtil.GetDocTestPath() + "option.user");
            Assert.AreEqual("Test Option Name", loadedOption.OptionName);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetFontFamily()
        {
            option.FontFamily = "test family";
            Assert.AreEqual("test family", option.GetFontFamily());
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTextBoxPosition()
        {
            Assert.AreEqual(Position.Centre, option.GetTextBoxPosition());
            option.TextBoxPosition = 4;
            Assert.AreEqual(Position.Left, option.GetTextBoxPosition());
            option.TextBoxPosition = 5;
            Assert.AreEqual(Position.Centre, option.GetTextBoxPosition());
            option.TextBoxPosition = 7;
            Assert.AreEqual(Position.BottomLeft, option.GetTextBoxPosition());
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTextBoxAlignment()
        {
            Assert.AreEqual(Alignment.Auto, option.GetTextBoxAlignment());
            option.TextBoxAlignment = 1;
            Assert.AreEqual(Alignment.Left, option.GetTextBoxAlignment());
        }
    }
}
