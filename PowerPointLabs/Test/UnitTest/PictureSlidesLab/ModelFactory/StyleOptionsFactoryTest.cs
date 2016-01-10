using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs;
using PowerPointLabs.PictureSlidesLab.ModelFactory;

namespace Test.UnitTest.PictureSlidesLab.ModelFactory
{
    [TestClass]
    public class StyleOptionsFactoryTest
    {
        [TestMethod]
        [TestCategory("UT")]
        public void TestGetAllVariationStyleOptions()
        {
            var allOptions = StyleOptionsFactory.GetAllStylesVariationOptions();
            Assert.IsTrue(allOptions.Count > 0);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetAllPreviewStyleOptions()
        {
            var allOptions = StyleOptionsFactory.GetAllStylesPreviewOptions();
            Assert.IsTrue(allOptions.Count > 0);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetDirectTextOptions()
        {
            var option = StyleOptionsFactory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameDirectText);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameDirectText, option.StyleName);

            var options = StyleOptionsFactory.GetStylesVariationOptions(
                TextCollection.PictureSlidesLabText.StyleNameDirectText);
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetBlurOptions()
        {
            var option = StyleOptionsFactory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameBlur);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameBlur, option.StyleName);
            Assert.IsTrue(option.IsUseBlurStyle);

            var options = StyleOptionsFactory.GetStylesVariationOptions(
                TextCollection.PictureSlidesLabText.StyleNameBlur);
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTextBoxOptions()
        {
            var option = StyleOptionsFactory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameTextBox);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameTextBox, option.StyleName);
            Assert.IsTrue(option.IsUseTextBoxStyle);

            var options = StyleOptionsFactory.GetStylesVariationOptions(
                TextCollection.PictureSlidesLabText.StyleNameTextBox);
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetBannerOptions()
        {
            var option = StyleOptionsFactory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameBanner);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameBanner, option.StyleName);
            Assert.IsTrue(option.IsUseBannerStyle);

            var options = StyleOptionsFactory.GetStylesVariationOptions(
                TextCollection.PictureSlidesLabText.StyleNameBanner);
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetSpecialEffectOptions()
        {
            var option = StyleOptionsFactory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameSpecialEffect);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameSpecialEffect, option.StyleName);
            Assert.IsTrue(option.IsUseSpecialEffectStyle);

            var options = StyleOptionsFactory.GetStylesVariationOptions(
                TextCollection.PictureSlidesLabText.StyleNameSpecialEffect);
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetOverlayOptions()
        {
            var option = StyleOptionsFactory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameOverlay);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameOverlay, option.StyleName);
            Assert.IsTrue(option.IsUseOverlayStyle);

            var options = StyleOptionsFactory.GetStylesVariationOptions(
                TextCollection.PictureSlidesLabText.StyleNameOverlay);
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetOutlineOptions()
        {
            var option = StyleOptionsFactory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameOutline);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameOutline, option.StyleName);
            Assert.IsTrue(option.IsUseOutlineStyle);

            var options = StyleOptionsFactory.GetStylesVariationOptions(
                TextCollection.PictureSlidesLabText.StyleNameOutline);
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetFrameOptions()
        {
            var option = StyleOptionsFactory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameFrame);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameFrame, option.StyleName);
            Assert.IsTrue(option.IsUseFrameStyle);

            var options = StyleOptionsFactory.GetStylesVariationOptions(
                TextCollection.PictureSlidesLabText.StyleNameFrame);
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetCircleOptions()
        {
            var option = StyleOptionsFactory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameCircle);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameCircle, option.StyleName);
            Assert.IsTrue(option.IsUseCircleStyle);

            var options = StyleOptionsFactory.GetStylesVariationOptions(
                TextCollection.PictureSlidesLabText.StyleNameCircle);
            Assert.AreEqual(8, options.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTriangleOptions()
        {
            var option = StyleOptionsFactory.GetStylesPreviewOption(
                TextCollection.PictureSlidesLabText.StyleNameTriangle);
            Assert.AreEqual(TextCollection.PictureSlidesLabText.StyleNameTriangle, option.StyleName);
            Assert.IsTrue(option.IsUseTriangleStyle);

            var options = StyleOptionsFactory.GetStylesVariationOptions(
                TextCollection.PictureSlidesLabText.StyleNameTriangle);
            Assert.AreEqual(8, options.Count);
        }
    }
}
