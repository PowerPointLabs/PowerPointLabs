using Microsoft.VisualStudio.TestTools.UnitTesting;
using Test.Util;

namespace Test.UnitTest
{
    [TestClass]
    public class FitToSlideUnitTest : BaseUnitTest
    {
        protected override string GetTestingSlideName()
        {
            return "FitToSlide.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void UT_FitToSlideUsualCases()
        {
            var slideWidth = Pres.PageSetup.SlideWidth;
            var slideHeight = Pres.PageSetup.SlideHeight;

            // Fit to width normal

            PpOperations.SelectSlide(4);
            var actualShape = PpOperations.SelectShape("pic")[1];
            PowerPointLabs.FitToSlide.FitToWidth(actualShape, slideWidth, slideHeight);

            PpOperations.SelectSlide(5);
            var expectedResultForFitToWidth = PpOperations.SelectShape("pic")[1];

            SlideUtil.IsSameShape(expectedResultForFitToWidth, actualShape);

            // Fit to height normal

            PowerPointLabs.FitToSlide.FitToHeight(actualShape, slideWidth, slideHeight);

            PpOperations.SelectSlide(6);
            var expectedResultForFitToHeight = PpOperations.SelectShape("pic")[1];

            SlideUtil.IsSameShape(expectedResultForFitToHeight, actualShape);

            // Fit to width rotated

            PpOperations.SelectSlide(8);
            actualShape = PpOperations.SelectShape("pic")[1];
            PowerPointLabs.FitToSlide.FitToWidth(actualShape, slideWidth, slideHeight);

            PpOperations.SelectSlide(9);
            expectedResultForFitToWidth = PpOperations.SelectShape("pic")[1];

            SlideUtil.IsSameShape(expectedResultForFitToWidth, actualShape);

            // Fit to height rotated

            PowerPointLabs.FitToSlide.FitToHeight(actualShape, slideWidth, slideHeight);

            PpOperations.SelectSlide(10);
            expectedResultForFitToHeight = PpOperations.SelectShape("pic")[1];

            SlideUtil.IsSameShape(expectedResultForFitToHeight, actualShape);
        }
    }
}
