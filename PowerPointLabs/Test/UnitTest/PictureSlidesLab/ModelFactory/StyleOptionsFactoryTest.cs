using System.Collections.Generic;
using System.Linq;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory;
using PowerPointLabs.TextCollection;

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
            List<List<StyleOption>> allOptions = _factory.GetAllStylesVariationOptions();
            Assert.IsTrue(allOptions.Count > 0);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetAllPreviewStyleOptions()
        {
            List<StyleOption> allOptions = _factory.GetAllStylesPreviewOptions();
            Assert.IsTrue(allOptions.Count > 0);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetDirectTextOptions()
        {
            StyleOption option = _factory.GetStylesPreviewOption(PictureSlidesLabText.StyleNameDirectText);
            Assert.AreEqual(PictureSlidesLabText.StyleNameDirectText, option.StyleName);

            List<StyleOption> options = GetOptions(PictureSlidesLabText.StyleNameDirectText);
            Assert.AreEqual(8, 
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseOverlayStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetBlurOptions()
        {
            StyleOption option = _factory.GetStylesPreviewOption(PictureSlidesLabText.StyleNameBlur);
            Assert.AreEqual(PictureSlidesLabText.StyleNameBlur, option.StyleName);
            Assert.IsTrue(option.IsUseBlurStyle);

            List<StyleOption> options = GetOptions(PictureSlidesLabText.StyleNameBlur);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseBlurStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTextBoxOptions()
        {
            StyleOption option = _factory.GetStylesPreviewOption(PictureSlidesLabText.StyleNameTextBox);
            Assert.AreEqual(PictureSlidesLabText.StyleNameTextBox, option.StyleName);
            Assert.IsTrue(option.IsUseTextBoxStyle);

            List<StyleOption> options = GetOptions(PictureSlidesLabText.StyleNameTextBox);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseTextBoxStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetBannerOptions()
        {
            StyleOption option = _factory.GetStylesPreviewOption(PictureSlidesLabText.StyleNameBanner);
            Assert.AreEqual(PictureSlidesLabText.StyleNameBanner, option.StyleName);
            Assert.IsTrue(option.IsUseBannerStyle);

            List<StyleOption> options = GetOptions(PictureSlidesLabText.StyleNameBanner);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseBannerStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetSpecialEffectOptions()
        {
            StyleOption option = _factory.GetStylesPreviewOption(PictureSlidesLabText.StyleNameSpecialEffect);
            Assert.AreEqual(PictureSlidesLabText.StyleNameSpecialEffect, option.StyleName);
            Assert.IsTrue(option.IsUseSpecialEffectStyle);

            List<StyleOption> options = GetOptions(PictureSlidesLabText.StyleNameSpecialEffect);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseSpecialEffectStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetOverlayOptions()
        {
            StyleOption option = _factory.GetStylesPreviewOption(PictureSlidesLabText.StyleNameOverlay);
            Assert.AreEqual(PictureSlidesLabText.StyleNameOverlay, option.StyleName);
            Assert.IsTrue(option.IsUseOverlayStyle);

            List<StyleOption> options = GetOptions(PictureSlidesLabText.StyleNameOverlay);
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
            StyleOption option = _factory.GetStylesPreviewOption(PictureSlidesLabText.StyleNameOutline);
            Assert.AreEqual(PictureSlidesLabText.StyleNameOutline, option.StyleName);
            Assert.IsTrue(option.IsUseOutlineStyle);

            List<StyleOption> options = GetOptions(PictureSlidesLabText.StyleNameOutline);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseOutlineStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetFrameOptions()
        {
            StyleOption option = _factory.GetStylesPreviewOption(PictureSlidesLabText.StyleNameFrame);
            Assert.AreEqual(PictureSlidesLabText.StyleNameFrame, option.StyleName);
            Assert.IsTrue(option.IsUseFrameStyle);

            List<StyleOption> options = GetOptions(PictureSlidesLabText.StyleNameFrame);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseFrameStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetCircleOptions()
        {
            StyleOption option = _factory.GetStylesPreviewOption(PictureSlidesLabText.StyleNameCircle);
            Assert.AreEqual(PictureSlidesLabText.StyleNameCircle, option.StyleName);
            Assert.IsTrue(option.IsUseCircleStyle);

            List<StyleOption> options = GetOptions(PictureSlidesLabText.StyleNameCircle);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseCircleStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTriangleOptions()
        {
            StyleOption option = _factory.GetStylesPreviewOption(PictureSlidesLabText.StyleNameTriangle);
            Assert.AreEqual(PictureSlidesLabText.StyleNameTriangle, option.StyleName);
            Assert.IsTrue(option.IsUseTriangleStyle);

            List<StyleOption> options = GetOptions(PictureSlidesLabText.StyleNameTriangle);
            Assert.AreEqual(8,
                GetExpectedCount(
                    GetOptionsProperty(options, "IsUseTriangleStyle"), true));
            Assert.AreEqual(8, options.Count);
        }

        private List<StyleOption> GetOptions(string styleName)
        {
            List<StyleOption> options = _factory.GetStylesVariationOptions(styleName);
            Dictionary<string, List<StyleVariant>> variants = new StyleVariantsFactory().GetVariants(styleName);

            for (int i = 0; i < options.Count; i++)
            {
                variants[variants.Keys.First()][i].Apply(options[i]);
            }

            return options;
        }

        private static List<object> GetOptionsProperty(List<StyleOption> options, string propertyName)
        {
            List<object> propList = new List<object>();
            foreach (StyleOption option in options)
            {
                System.Type type = option.GetType();
                System.Reflection.PropertyInfo prop = type.GetProperty(propertyName);
                object propValue = prop.GetValue(option, null);
                propList.Add(propValue);
            }
            return propList;
        }

        private static int GetExpectedCount(List<object> list, object expected)
        {
            int result = 0;
            foreach (object item in list)
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
