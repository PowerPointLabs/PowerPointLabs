using System.Collections.Generic;
using System.Linq;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory;
using PowerPointLabs.TextCollection;

namespace Test.UnitTest.PictureSlidesLab.ModelFactory
{
    [TestClass]
    public class StyleVariantsFactoryTest
    {
        private StyleVariantsFactory _variantsFactory;
        private StyleOptionsFactory _optionsFactory;

        [TestInitialize]
        public void Init()
        {
            _variantsFactory = new StyleVariantsFactory();
            _optionsFactory = new StyleOptionsFactory();
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetDirectTextVariants()
        {
            VerifyVariants(PictureSlidesLabText.StyleNameDirectText);
            VerifyVariants2(PictureSlidesLabText.StyleNameDirectText);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetBlurVariants()
        {
            VerifyVariants(PictureSlidesLabText.StyleNameBlur);
            VerifyVariants2(PictureSlidesLabText.StyleNameBlur);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTextBoxVariants()
        {
            VerifyVariants(PictureSlidesLabText.StyleNameTextBox);
            VerifyVariants2(PictureSlidesLabText.StyleNameTextBox);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetBannerVariants()
        {
            VerifyVariants(PictureSlidesLabText.StyleNameBanner);
            VerifyVariants2(PictureSlidesLabText.StyleNameBanner);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetSpecialEffectVariants()
        {
            VerifyVariants(PictureSlidesLabText.StyleNameSpecialEffect);
            VerifyVariants2(PictureSlidesLabText.StyleNameSpecialEffect);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetOverlayVariants()
        {
            VerifyVariants(PictureSlidesLabText.StyleNameOverlay);
            VerifyVariants2(PictureSlidesLabText.StyleNameOverlay);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetOutlineVariants()
        {
            VerifyVariants(PictureSlidesLabText.StyleNameOutline);
            VerifyVariants2(PictureSlidesLabText.StyleNameOutline);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetFrameVariants()
        {
            VerifyVariants(PictureSlidesLabText.StyleNameFrame);
            VerifyVariants2(PictureSlidesLabText.StyleNameFrame);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetCircleVariants()
        {
            VerifyVariants(PictureSlidesLabText.StyleNameCircle);
            VerifyVariants2(PictureSlidesLabText.StyleNameCircle);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTriangleVariants()
        {
            VerifyVariants(PictureSlidesLabText.StyleNameTriangle);
            VerifyVariants2(PictureSlidesLabText.StyleNameTriangle);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void VerifyVariantsKeysCount()
        {
            IEnumerable<PowerPointLabs.PictureSlidesLab.ModelFactory.Variants.Interface.IStyleVariants> allVariants =
                _variantsFactory.GetAllStyleVariants();
            foreach (PowerPointLabs.PictureSlidesLab.ModelFactory.Variants.Interface.IStyleVariants styleVariants in allVariants)
            {
                foreach (IEnumerable<StyleVariant> styleVariant in styleVariants.GetVariantsForStyle().Values)
                {
                    int expectedKeyCount = styleVariant.First().GetVariants().Keys.Count;
                    foreach (StyleVariant variant in styleVariant)
                    {
                        Assert.AreEqual(expectedKeyCount, variant.GetVariants().Keys.Count,
                            "Keys Count should be same for a VariantWorker.");
                    }
                }
            }
        }

        private void VerifyVariants(string styleName)
        {
            Dictionary<string, List<StyleVariant>> variants =
                _variantsFactory.GetVariants(styleName);
            StyleOption option =
                _optionsFactory.GetStylesPreviewOption(styleName);

            int numberOfNoEffectVariant = 0;
            foreach (string key in variants.Keys)
            {
                if (key == PictureSlidesLabText.VariantCategoryFontFamily)
                    continue;

                List<StyleVariant> variant = variants[key];
                Assert.AreEqual(8, variant.Count,
                    "Each variant/category/aspect/dimension should have 8 variations");
                foreach (StyleVariant styleVariants in variant)
                {
                    if (styleVariants.IsNoEffect(option))
                    {
                        numberOfNoEffectVariant++;
                    }
                }
            }
            Assert.AreEqual(variants.Values.Count - 1, numberOfNoEffectVariant,
                "In order to swap no effect variant with the style correctly, it is assumed that " +
                "number of no effect variant should be equal to number of variants/category/aspect/dimension. " +
                "Please modify a variation to have no effect on the style. Ref: issue #802.");
        }

        private void VerifyVariants2(string styleName)
        {
            Dictionary<string, List<StyleVariant>> variants =
                _variantsFactory.GetVariants(styleName);
            List<StyleOption> options =
                _optionsFactory.GetStylesVariationOptions(styleName);

            for (int i = 0; i < options.Count; i++)
            {
                variants[variants.Keys.First()][i].Apply(options[i]);
            }

            int numberOfNoEffectVariant = 0;
            foreach (string key in variants.Keys)
            {
                if (key == PictureSlidesLabText.VariantCategoryFontFamily)
                    continue;

                List<StyleVariant> variant = variants[key];
                foreach (StyleVariant styleVariants in variant)
                {
                    foreach (StyleOption option in options)
                    {
                        if (styleVariants.IsNoEffect(option))
                        {
                            numberOfNoEffectVariant++;
                        }
                    }
                }
            }
            Assert.AreEqual((variants.Values.Count - 1) * options.Count, numberOfNoEffectVariant,
                "In order to swap no effect variant with the style correctly, it is assumed that " +
                "number of no effect variant should be equal to number of variants/category/aspect/dimension. " +
                "Please modify a variation to have no effect on the style. Ref: issue #802.");
        }
    }
}
