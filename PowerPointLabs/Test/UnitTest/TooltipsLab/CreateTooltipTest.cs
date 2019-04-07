using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.TooltipsLab;
using PowerPointLabs.Models;

using Test.Util;
using Microsoft.Office.Interop.PowerPoint;

namespace Test.UnitTest.TooltipsLab
{
    [TestClass]
    public class CreateTooltipTest : BaseUnitTest
    {
        protected override string GetTestingSlideName()
        {
            return "TooltipsLab\\CreateTooltip.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void CreateTooltip()
        {
            TestCreateTooltip_NoneSelected();
            TestCreateTooltip_ShapeSelected();
            TestCreateTooltip_ExistingTooltip();
        }

        private const int CreateTooltipNoneSelectedTestSlideNo = 4;
        private const int CreateTooltipNoneSelectedExpectedSlideNo = 5;
        private const int CreateTooltipShapeSelectedTestSlideNo = 7;
        private const int CreateTooltipShapeSelectedExpectedSlideNo = 8;
        private const int CreateTooltipExistingTooltipTestSlideNo = 10;
        private const int CreateTooltipExistingTooltipExpectedSlideNo = 11;
        private const string CreateTooltipShapeToSelectName = "SelectMe";

        private void TestCreateTooltip_NoneSelected()
        {
            PpOperations.SelectSlide(CreateTooltipNoneSelectedTestSlideNo);
            CreateTooltipAndCompare(CreateTooltipNoneSelectedTestSlideNo, CreateTooltipNoneSelectedExpectedSlideNo);
        }

        private void TestCreateTooltip_ShapeSelected()
        {
            PpOperations.SelectSlide(CreateTooltipShapeSelectedTestSlideNo);
            PpOperations.SelectShape(CreateTooltipShapeToSelectName);
            CreateTooltipAndCompare(CreateTooltipShapeSelectedTestSlideNo, CreateTooltipShapeSelectedExpectedSlideNo);
        }
       
        private void TestCreateTooltip_ExistingTooltip()
        {
            PpOperations.SelectSlide(CreateTooltipExistingTooltipTestSlideNo);
            CreateTooltipAndCompare(CreateTooltipExistingTooltipTestSlideNo, CreateTooltipExistingTooltipExpectedSlideNo);
        }


        private void CreateTooltipAndCompare(int testSlideNo, int expectedSlideNo)
        {
            PowerPointSlide currentSlide = PowerPointSlide.FromSlideFactory(PpOperations.GetCurrentSlide());
            Shape triggerShape = PowerPointLabs.TooltipsLab.CreateTooltip.GenerateTriggerShape(currentSlide);
            Shape callout = PowerPointLabs.TooltipsLab.CreateTooltip.GenerateCalloutWithReferenceTriggerShape(currentSlide, triggerShape);
            ConvertToTooltip.AddTriggerAnimation(currentSlide, triggerShape, callout);
            AssertIsSame(testSlideNo, expectedSlideNo);
        }

        private void AssertIsSame(int actualSlideIndex, int expectedSlideIndex)
        {
            Slide actualSlide = PpOperations.SelectSlide(actualSlideIndex);
            Slide expectedSlide = PpOperations.SelectSlide(expectedSlideIndex);
            SlideUtil.IsSameLooking(expectedSlide, actualSlide);
            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
        }
    }
}
