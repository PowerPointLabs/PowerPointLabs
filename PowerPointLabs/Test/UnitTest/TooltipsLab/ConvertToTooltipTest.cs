﻿using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.TooltipsLab;
using PowerPointLabs.Models;

using Test.Util;
using Microsoft.Office.Interop.PowerPoint;

namespace Test.UnitTest.TooltipsLab
{
    [TestClass]
    public class ConvertToTooltipTest : BaseUnitTest
    {
        private const int ConvertShapesToTooltipOneShapeTestSlideNo = 4;
        private const int ConvertShapesToTooltipOneShapeExpectedSlideNo = 5;
        private const int ConvertShapesToTooltipTwoShapesTestSlideNo = 7;
        private const int ConvertShapesToTooltipTwoShapesExpectedSlideNo = 8;
        private const int ConvertShapesToTooltipThreeShapesTestSlideNo = 10;
        private const int ConvertShapesToTooltipThreeShapesExpectedSlideNo = 11;

        private const string TriggerShapeName = "Trigger";
        private const string CalloutShapeName = "Callout";
        private const string Callout2ShapeName = "Callout2";

        protected override string GetTestingSlideName()
        {
            return "TooltipsLab\\ConvertToTooltip.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void ConvertShapesToTooltip()
        {
            TestConvertShapesToTooltip_TwoShapes();
            TestConvertShapesToTooltip_ThreeShapes();
        }

        private void TestConvertShapesToTooltip_OneShape()
        {
            Slide slide = PpOperations.SelectSlide(ConvertShapesToTooltipOneShapeTestSlideNo);
            string[] shapeNames = { TriggerShapeName };
            ConvertShapesToTooltipAndCompare(shapeNames, ConvertShapesToTooltipOneShapeTestSlideNo, ConvertShapesToTooltipOneShapeExpectedSlideNo, false);
        }

        private void TestConvertShapesToTooltip_TwoShapes()
        {
            Slide slide = PpOperations.SelectSlide(ConvertShapesToTooltipTwoShapesTestSlideNo);
            string[] shapeNames = { TriggerShapeName, CalloutShapeName };
            ConvertShapesToTooltipAndCompare(shapeNames, ConvertShapesToTooltipTwoShapesTestSlideNo, ConvertShapesToTooltipTwoShapesExpectedSlideNo, true);
        }

        private void TestConvertShapesToTooltip_ThreeShapes()
        {
            PpOperations.SelectSlide(ConvertShapesToTooltipTwoShapesTestSlideNo);
            string[] shapeNames = { TriggerShapeName, CalloutShapeName, Callout2ShapeName };
            ConvertShapesToTooltipAndCompare(shapeNames, ConvertShapesToTooltipTwoShapesTestSlideNo, ConvertShapesToTooltipTwoShapesExpectedSlideNo, true);
        }

        private void ConvertShapesToTooltipAndCompare(string[] shapeNames, int testSlideNo, int expectedSlideNo, bool isSuccessful)
        {
            PowerPointSlide currentSlide = PowerPointSlide.FromSlideFactory(PpOperations.GetCurrentSlide());
            PpOperations.SelectShapes(shapeNames);
            Selection selection = PpOperations.GetCurrentSelection();
            Assert.Equals(ConvertToTooltip.AddTriggerAnimation(currentSlide, selection), isSuccessful);
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
