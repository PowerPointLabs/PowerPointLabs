using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;

namespace Test.UnitTest
{
    [TestClass]
    public class FitToSlideTest : BaseUnitTest
    {
        protected override string GetTestingSlideName()
        {
            return "FitToSlide.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void FitToWidth()
        {
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;

            PpOperations.SelectSlide(4);
            var actualShape = PpOperations.SelectShape("pic")[1];
            PowerPointLabs.FitToSlide.FitToWidth(actualShape, slideWidth, slideHeight);

            PpOperations.SelectSlide(5);
            var expectedResultForFitToWidth = PpOperations.SelectShape("pic")[1];

            SlideUtil.IsSameShape(expectedResultForFitToWidth, actualShape);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void FitToHeight()
        {
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;

            PpOperations.SelectSlide(4);
            var actualShape = PpOperations.SelectShape("pic")[1];

            PowerPointLabs.FitToSlide.FitToHeight(actualShape, slideWidth, slideHeight);

            PpOperations.SelectSlide(6);
            var expectedResultForFitToHeight = PpOperations.SelectShape("pic")[1];

            SlideUtil.IsSameShape(expectedResultForFitToHeight, actualShape);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void FitToWidthRotated()
        {
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;

            PpOperations.SelectSlide(8);
            var actualShape = PpOperations.SelectShape("pic")[1];
            PowerPointLabs.FitToSlide.FitToWidth(actualShape, slideWidth, slideHeight);

            PpOperations.SelectSlide(9);
            var expectedResultForFitToWidth = PpOperations.SelectShape("pic")[1];

            SlideUtil.IsSameShape(expectedResultForFitToWidth, actualShape);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void FitToHeightRotated()
        {
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;

            PpOperations.SelectSlide(8);
            var actualShape = PpOperations.SelectShape("pic")[1];
            PowerPointLabs.FitToSlide.FitToHeight(actualShape, slideWidth, slideHeight);

            PpOperations.SelectSlide(10);
            var expectedResultForFitToHeight = PpOperations.SelectShape("pic")[1];

            SlideUtil.IsSameShape(expectedResultForFitToHeight, actualShape);
        }
    }
}
