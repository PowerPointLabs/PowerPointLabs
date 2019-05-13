using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.Models;
using PowerPointLabs.TooltipsLab;

using Test.Util;

namespace Test.UnitTest.TooltipsLab
{
    [TestClass]
    public class CreateTooltipTest : BaseUnitTest
    {
        private const int CreateTooltipNoneSelectedTestSlideNo = 4;
        private const int CreateTooltipNoneSelectedExpectedSlideNo = 5;
        private const int CreateTooltipShapeSelectedTestSlideNo = 7;
        private const int CreateTooltipShapeSelectedExpectedSlideNo = 8;
        private const int CreateTooltipExistingTooltipTestSlideNo = 10;
        private const int CreateTooltipExistingTooltipExpectedSlideNo = 11;
        private const int CreateTooltipMultipleTooltipsTestSlideNo = 13;
        private const int CreateTooltipMultipleTooltipsExpectedSlideNo = 14;

        private const string CreateTooltipShapeToSelectName = "SelectMe";
        private const int NumberOfTooltips = 5;

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
            TestCreateTooltip_MultipleTooltips();
        }

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

        private void TestCreateTooltip_MultipleTooltips()
        {
            PpOperations.SelectSlide(CreateTooltipMultipleTooltipsTestSlideNo);
            for (int i = 0; i < NumberOfTooltips; i++)
            {
                CreateTooltipOnSlide(CreateTooltipMultipleTooltipsTestSlideNo);
            }
            AssertIsSame(CreateTooltipMultipleTooltipsTestSlideNo, CreateTooltipMultipleTooltipsExpectedSlideNo);
        }

        private void CreateTooltipAndCompare(int testSlideNo, int expectedSlideNo)
        {
            CreateTooltipOnSlide(testSlideNo);
            AssertIsSame(testSlideNo, expectedSlideNo);
        }

        private void CreateTooltipOnSlide(int slideNo)
        {
            PowerPointSlide currentSlide = PowerPointSlide.FromSlideFactory(PpOperations.GetCurrentSlide());
            Shape triggerShape = PowerPointLabs.TooltipsLab.CreateTooltip.GenerateTriggerShape(currentSlide);
            Shape callout = PowerPointLabs.TooltipsLab.CreateTooltip.GenerateCalloutWithReferenceTriggerShape(currentSlide, triggerShape);
            ConvertToTooltip.AddTriggerAnimation(currentSlide, triggerShape, callout);
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
