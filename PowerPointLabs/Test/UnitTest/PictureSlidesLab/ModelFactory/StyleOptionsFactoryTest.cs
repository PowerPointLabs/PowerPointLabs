using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory;

namespace Test.UnitTest.PictureSlidesLab.ModelFactory
{
    [TestClass]
    public class StyleOptionsFactoryTest
    {
        private StyleOptionsFactory _factory;

        [TestInitialize]
        public void Init()
        {
            _factory = new StyleOptionsFactory();
        }
        
        [TestMethod]
        [TestCategory("UT")]
        public void TestGetAllVariationStyleOptions()
        {
            var allOptions = _factory.GetAllStylesVariationOptions();
            Assert.IsTrue(allOptions.Count > 0);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetAllPreviewStyleOptions()
        {
            var allOptions = _factory.GetAllStylesPreviewOptions();
            Assert.IsTrue(allOptions.Count > 0);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetDirectTextOptions()
        {
            var option = _factory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameDirectText);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameDirectText, option.StyleName);

            var options = GetOptions(TextCollection.PictureSlidesLabText.StyleNameDirectText);
            Assert.AreEqual(8, 
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseOverlayStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetBlurOptions()
        {
            var option = _factory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameBlur);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameBlur, option.StyleName);
            Assert.IsTrue(option.IsUseBlurStyle);
            
            var options = GetOptions(TextCollection.PictureSlidesLabText.StyleNameBlur);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseBlurStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTextBoxOptions()
        {
            var option = _factory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameTextBox);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameTextBox, option.StyleName);
            Assert.IsTrue(option.IsUseTextBoxStyle);

            var options = GetOptions(
                TextCollection.PictureSlidesLabText.StyleNameTextBox);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseTextBoxStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetBannerOptions()
        {
            var option = _factory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameBanner);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameBanner, option.StyleName);
            Assert.IsTrue(option.IsUseBannerStyle);

            var options = GetOptions(
                TextCollection.PictureSlidesLabText.StyleNameBanner);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseBannerStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetSpecialEffectOptions()
        {
            var option = _factory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameSpecialEffect);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameSpecialEffect, option.StyleName);
            Assert.IsTrue(option.IsUseSpecialEffectStyle);

            var options = GetOptions(
                TextCollection.PictureSlidesLabText.StyleNameSpecialEffect);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseSpecialEffectStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetOverlayOptions()
        {
            var option = _factory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameOverlay);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameOverlay, option.StyleName);
            Assert.IsTrue(option.IsUseOverlayStyle);

            var options = GetOptions(
                TextCollection.PictureSlidesLabText.StyleNameOverlay);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseOverlayStyle"), true));
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseSpecialEffectStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetOutlineOptions()
        {
            var option = _factory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameOutline);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameOutline, option.StyleName);
            Assert.IsTrue(option.IsUseOutlineStyle);

            var options = GetOptions(
                TextCollection.PictureSlidesLabText.StyleNameOutline);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseOutlineStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetFrameOptions()
        {
            var option = _factory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameFrame);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameFrame, option.StyleName);
            Assert.IsTrue(option.IsUseFrameStyle);

            var options = GetOptions(
                TextCollection.PictureSlidesLabText.StyleNameFrame);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseFrameStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetCircleOptions()
        {
            var option = _factory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameCircle);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameCircle, option.StyleName);
            Assert.IsTrue(option.IsUseCircleStyle);

            var options = GetOptions(
                TextCollection.PictureSlidesLabText.StyleNameCircle);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseCircleStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTriangleOptions()
        {
            var option = _factory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameTriangle);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameTriangle, option.StyleName);
            Assert.IsTrue(option.IsUseTriangleStyle);

            var options = GetOptions(
                TextCollection.PictureSlidesLabText.StyleNameTriangle);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseTriangleStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        private List<StyleOption> GetOptions(string styleName)
        {
            var options = _factory.GetStylesVariationOptions(styleName);
            var variants = new StyleVariantsFactory().GetVariants(styleName);

            for (var i = 0; i < options.Count; i++)
            {
                variants[variants.Keys.First()][i].Apply(options[i]);
            }

            return options;
        }

        private static List<object> GetOptionsProperty(List<StyleOption> options, string propertyName)
        {
            var propList = new List<object>();
            foreach (var option in options)
            {
                var type = option.GetType();
                var prop = type.GetProperty(propertyName);
                var propValue = prop.GetValue(option, null);
                propList.Add(propValue);
            }
            return propList;
        }

        private static int GetExpectedCount(List<object> list, object expected)
        {
            var result = 0;
            foreach (var item in list)
            {
                if (item.Equals(expected))
                {
                    result++;
                }
            }
            return result;
        }
    }
}
