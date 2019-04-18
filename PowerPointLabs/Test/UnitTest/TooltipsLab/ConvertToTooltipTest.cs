using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.TooltipsLab;
using PowerPointLabs.Models;

using Test.Util;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.TextCollection;

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
        private const int ConvertShapesToTooltipTenShapesTestSlideNo = 13;
        private const int ConvertShapesToTooltipTenShapesExpectedSlideNo = 14;

        private const string TriggerShapeName = "Trigger";
        private const string CalloutShapeName = "Callout";
        private const string Callout2ShapeName = "Callout 2";
        private const string Callout3ShapeName = "Callout 3";
        private const string Callout4ShapeName = "Callout 4";
        private const string Callout5ShapeName = "Callout 5";
        private const string Callout6ShapeName = "Callout 6";
        private const string Callout7ShapeName = "Callout 7";
        private const string Callout8ShapeName = "Callout 8";
        private const string Callout9ShapeName = "Callout 9";

        protected override string GetTestingSlideName()
        {
            return "TooltipsLab\\ConvertToTooltip.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void ConvertShapesToTooltip()
        {
            TestConvertShapesToTooltip_OneShape();
            TestConvertShapesToTooltip_TwoShapes();
            TestConvertShapesToTooltip_ThreeShapes();
            TestConvertShapesToTooltip_TenShapes();
        }

        private void TestConvertShapesToTooltip_OneShape()
        {
            string[] shapeNames = { TriggerShapeName };
            ConvertShapesToTooltipAndCompare(shapeNames, ConvertShapesToTooltipOneShapeTestSlideNo, ConvertShapesToTooltipOneShapeExpectedSlideNo, false);
        }

        private void TestConvertShapesToTooltip_TwoShapes()
        {
            string[] shapeNames = { TriggerShapeName, CalloutShapeName };
            ConvertShapesToTooltipAndCompare(shapeNames, ConvertShapesToTooltipTwoShapesTestSlideNo, ConvertShapesToTooltipTwoShapesExpectedSlideNo, true);
        }

        private void TestConvertShapesToTooltip_ThreeShapes()
        {
            string[] shapeNames = { TriggerShapeName, CalloutShapeName, Callout2ShapeName };
            ConvertShapesToTooltipAndCompare(shapeNames, ConvertShapesToTooltipThreeShapesTestSlideNo, ConvertShapesToTooltipThreeShapesExpectedSlideNo, true);
        }

        private void TestConvertShapesToTooltip_TenShapes()
        {
            string[] shapeNames = { TriggerShapeName, CalloutShapeName, Callout2ShapeName, Callout3ShapeName,
                Callout4ShapeName, Callout5ShapeName, Callout6ShapeName,
                Callout7ShapeName, Callout8ShapeName, Callout9ShapeName};
            ConvertShapesToTooltipAndCompare(shapeNames, ConvertShapesToTooltipTenShapesTestSlideNo, ConvertShapesToTooltipTenShapesExpectedSlideNo, true);
        }

        private void ConvertShapesToTooltipAndCompare(string[] shapeNames, int testSlideNo, int expectedSlideNo, bool isSuccessful)
        {
            Slide slide = PpOperations.SelectSlide(testSlideNo);
            PowerPointSlide currentSlide = PowerPointSlide.FromSlideFactory(slide);
            ShapeRange selectedShapes = PpOperations.SelectShapes(shapeNames);
            if (isSuccessful)
            {
                Assert.AreEqual(ConvertToTooltip.AddTriggerAnimation(currentSlide, selectedShapes), isSuccessful);
            }
            else
            {
                
                MessageBoxUtil.ExpectMessageBoxWillPopUp(
                                            TooltipsLabText.ErrorTooltipsDialogTitle,
                                            TooltipsLabText.ErrorLessThanTwoShapesSelected,
                                            () => { ConvertToTooltip.AddTriggerAnimation(currentSlide, selectedShapes);  });
            }
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
