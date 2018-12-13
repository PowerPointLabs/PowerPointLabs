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
            float slideWidth = Pres.PageSetup.SlideWidth;
            float slideHeight = Pres.PageSetup.SlideHeight;

            PpOperations.SelectSlide(4);
            Microsoft.Office.Interop.PowerPoint.Shape actualShape = PpOperations.SelectShape("pic")[1];
            PowerPointLabs.FitToSlide.FitToWidth(actualShape, slideWidth, slideHeight);

            PpOperations.SelectSlide(5);
            Microsoft.Office.Interop.PowerPoint.Shape expectedResultForFitToWidth = PpOperations.SelectShape("pic")[1];

            SlideUtil.IsSameShape(expectedResultForFitToWidth, actualShape);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void FitToHeight()
        {
            float slideWidth = Pres.PageSetup.SlideWidth;
            float slideHeight = Pres.PageSetup.SlideHeight;

            PpOperations.SelectSlide(4);
            Microsoft.Office.Interop.PowerPoint.Shape actualShape = PpOperations.SelectShape("pic")[1];

            PowerPointLabs.FitToSlide.FitToHeight(actualShape, slideWidth, slideHeight);

            PpOperations.SelectSlide(6);
            Microsoft.Office.Interop.PowerPoint.Shape expectedResultForFitToHeight = PpOperations.SelectShape("pic")[1];

            SlideUtil.IsSameShape(expectedResultForFitToHeight, actualShape);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void FitToWidthRotated()
        {
            float slideWidth = Pres.PageSetup.SlideWidth;
            float slideHeight = Pres.PageSetup.SlideHeight;

            PpOperations.SelectSlide(8);
            Microsoft.Office.Interop.PowerPoint.Shape actualShape = PpOperations.SelectShape("pic")[1];
            PowerPointLabs.FitToSlide.FitToWidth(actualShape, slideWidth, slideHeight);

            PpOperations.SelectSlide(9);
            Microsoft.Office.Interop.PowerPoint.Shape expectedResultForFitToWidth = PpOperations.SelectShape("pic")[1];

            SlideUtil.IsSameShape(expectedResultForFitToWidth, actualShape);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void FitToHeightRotated()
        {
            float slideWidth = Pres.PageSetup.SlideWidth;
            float slideHeight = Pres.PageSetup.SlideHeight;

            PpOperations.SelectSlide(8);
            Microsoft.Office.Interop.PowerPoint.Shape actualShape = PpOperations.SelectShape("pic")[1];
            PowerPointLabs.FitToSlide.FitToHeight(actualShape, slideWidth, slideHeight);

            PpOperations.SelectSlide(10);
            Microsoft.Office.Interop.PowerPoint.Shape expectedResultForFitToHeight = PpOperations.SelectShape("pic")[1];

            SlideUtil.IsSameShape(expectedResultForFitToHeight, actualShape);
        }
    }
}
