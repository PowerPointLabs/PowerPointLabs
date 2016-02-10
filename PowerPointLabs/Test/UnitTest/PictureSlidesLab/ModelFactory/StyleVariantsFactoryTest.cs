using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory;

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
            VerifyVariants(TextCollection.PictureSlidesLabText.StyleNameDirectText);
            VerifyVariants2(TextCollection.PictureSlidesLabText.StyleNameDirectText);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetBlurVariants()
        {
            VerifyVariants(TextCollection.PictureSlidesLabText.StyleNameBlur);
            VerifyVariants2(TextCollection.PictureSlidesLabText.StyleNameBlur);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTextBoxVariants()
        {
            VerifyVariants(TextCollection.PictureSlidesLabText.StyleNameTextBox);
            VerifyVariants2(TextCollection.PictureSlidesLabText.StyleNameTextBox);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetBannerVariants()
        {
            VerifyVariants(TextCollection.PictureSlidesLabText.StyleNameBanner);
            VerifyVariants2(TextCollection.PictureSlidesLabText.StyleNameBanner);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetSpecialEffectVariants()
        {
            VerifyVariants(TextCollection.PictureSlidesLabText.StyleNameSpecialEffect);
            VerifyVariants2(TextCollection.PictureSlidesLabText.StyleNameSpecialEffect);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetOverlayVariants()
        {
            VerifyVariants(TextCollection.PictureSlidesLabText.StyleNameOverlay);
            VerifyVariants2(TextCollection.PictureSlidesLabText.StyleNameOverlay);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetOutlineVariants()
        {
            VerifyVariants(TextCollection.PictureSlidesLabText.StyleNameOutline);
            VerifyVariants2(TextCollection.PictureSlidesLabText.StyleNameOutline);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetFrameVariants()
        {
            VerifyVariants(TextCollection.PictureSlidesLabText.StyleNameFrame);
            VerifyVariants2(TextCollection.PictureSlidesLabText.StyleNameFrame);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetCircleVariants()
        {
            VerifyVariants(TextCollection.PictureSlidesLabText.StyleNameCircle);
            VerifyVariants2(TextCollection.PictureSlidesLabText.StyleNameCircle);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTriangleVariants()
        {
            VerifyVariants(TextCollection.PictureSlidesLabText.StyleNameTriangle);
            VerifyVariants2(TextCollection.PictureSlidesLabText.StyleNameTriangle);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void VerifyVariantsKeysCount()
        {
            var allVariants =
                _variantsFactory.GetAllStyleVariants();
            foreach (var styleVariants in allVariants)
            {
                foreach (IEnumerable<StyleVariant> styleVariant in styleVariants.GetVariantsForStyle().Values)
                {
                    var expectedKeyCount = styleVariant.First().GetVariants().Keys.Count;
                    foreach (var variant in styleVariant)
                    {
                        Assert.AreEqual(expectedKeyCount, variant.GetVariants().Keys.Count,
                            "Keys Count should be same for a VariantWorker.");
                    }
                }
            }
        }

        private void VerifyVariants(string styleName)
        {
            var variants =
                _variantsFactory.GetVariants(styleName);
            var option =
                _optionsFactory.GetStylesPreviewOption(styleName);

            var numberOfNoEffectVariant = 0;
            foreach (var variant in variants.Values)
            {
                Assert.AreEqual(8, variant.Count,
                    "Each variant/category/aspect/dimension should have 8 variations");
                foreach (var styleVariants in variant)
                {
                    if (styleVariants.IsNoEffect(option))
                    {
                        numberOfNoEffectVariant++;
                    }
                }
            }
            Assert.AreEqual(variants.Values.Count, numberOfNoEffectVariant,
                "In order to swap no effect variant with the style correctly, it is assumed that " +
                "number of no effect variant should be equal to number of variants/category/aspect/dimension. " +
                "Please modify a variation to have no effect on the style. Ref: issue #802.");
        }

        private void VerifyVariants2(string styleName)
        {
            var variants =
                _variantsFactory.GetVariants(styleName);
            var options =
                _optionsFactory.GetStylesVariationOptions(styleName);

            for (var i = 0; i < options.Count; i++)
            {
                variants[variants.Keys.First()][i].Apply(options[i]);
            }

            var numberOfNoEffectVariant = 0;
            foreach (var variant in variants.Values)
            {
                foreach (var styleVariants in variant)
                {
                    foreach (var option in options)
                    {
                        if (styleVariants.IsNoEffect(option))
                        {
                            numberOfNoEffectVariant++;
                        }
                    }
                }
            }
            Assert.AreEqual(variants.Values.Count * options.Count, numberOfNoEffectVariant,
                "In order to swap no effect variant with the style correctly, it is assumed that " +
                "number of no effect variant should be equal to number of variants/category/aspect/dimension. " +
                "Please modify a variation to have no effect on the style. Ref: issue #802.");
        }
    }
}
