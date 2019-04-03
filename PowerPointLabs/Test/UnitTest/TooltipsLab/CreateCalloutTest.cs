using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.TooltipsLab;
using PowerPointLabs.Models;

using Test.Util;
using Microsoft.Office.Interop.PowerPoint;

namespace Test.UnitTest.TooltipsLab
{
    [TestClass]
    public class CreateCalloutTest : BaseUnitTest
    {
        protected override string GetTestingSlideName()
        {
            return "TooltipsLab\\CreateCallout.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void CreateCallout()
        {
            TestCreateCallout_NormalShape();
            TestCreateCallout_TriggerShape();
            TestCreateCallout_CalloutShape();
            TestCreateCallout_MultipleNormalShapes();
            TestCreateCallout_MultipleAllTypeShapes();
        }

        private const int CreateCalloutNormalShapeTestSlideNo = 4;
        private const int CreateCalloutNormalShapeExpectedSlideNo = 5;
        private const int CreateCalloutTriggerShapeTestSlideNo = 7;
        private const int CreateCalloutTriggerShapeExpectedSlideNo = 8;
        private const int CreateCalloutCalloutShapeTestSlideNo = 10;
        private const int CreateCalloutCalloutShapeExpectedSlideNo = 11;
        private const int CreateCalloutMultipleNormalShapesTestSlideNo = 13;
        private const int CreateCalloutMultipleNormalShapesExpectedSlideNo = 14;
        private const int CreateCalloutMultipleAllTypeShapesTestSlideNo = 16;
        private const int CreateCalloutMultipleAllTypeShapesExpectedSlideNo = 17;
        private const string NormalShapeName = "normalShape";
        private const string TriggerShapeName = "existingTriggerShape";
        private const string CalloutShapeName = "existingCalloutShape";
        private const string MultipleShapePrefix = "select";

        private void TestCreateCallout_NormalShape()
        {
            PpOperations.SelectSlide(CreateCalloutNormalShapeTestSlideNo);
            ShapeRange selectedShapeRange = PpOperations.SelectShape(NormalShapeName);
            CreateCalloutAndCompare(selectedShapeRange[1], CreateCalloutNormalShapeTestSlideNo, CreateCalloutNormalShapeExpectedSlideNo);
        }

        private void TestCreateCallout_TriggerShape()
        {
            PpOperations.SelectSlide(CreateCalloutTriggerShapeTestSlideNo);
            ShapeRange selectedShapeRange = PpOperations.SelectShape(TriggerShapeName);
            CreateCalloutAndCompare(selectedShapeRange[1], CreateCalloutTriggerShapeTestSlideNo, CreateCalloutTriggerShapeExpectedSlideNo);
        }

        private void TestCreateCallout_CalloutShape()
        {
            PpOperations.SelectSlide(CreateCalloutCalloutShapeTestSlideNo);
            ShapeRange selectedShapeRange = PpOperations.SelectShape(CalloutShapeName);
            CreateCalloutAndCompare(selectedShapeRange[1], CreateCalloutCalloutShapeTestSlideNo, CreateCalloutCalloutShapeExpectedSlideNo);
        }

        private void TestCreateCallout_MultipleNormalShapes()
        {
            PpOperations.SelectSlide(CreateCalloutMultipleNormalShapesTestSlideNo);
            ShapeRange selectedShapeRange = PpOperations.SelectShapesByPrefix(NormalShapeName);
            CreateMultipleCalloutAndCompare(selectedShapeRange, CreateCalloutMultipleNormalShapesTestSlideNo, CreateCalloutMultipleNormalShapesExpectedSlideNo);
        }

        private void TestCreateCallout_MultipleAllTypeShapes()
        {
            PpOperations.SelectSlide(CreateCalloutMultipleAllTypeShapesTestSlideNo);
            ShapeRange selectedShapeRange = PpOperations.SelectShapesByPrefix(MultipleShapePrefix);
            CreateMultipleCalloutAndCompare(selectedShapeRange, CreateCalloutMultipleAllTypeShapesTestSlideNo, CreateCalloutMultipleAllTypeShapesExpectedSlideNo);
        }


        private void CreateCalloutAndCompare(Shape selectedShape, int testSlideNo, int expectedSlideNo)
        {
            PowerPointSlide currentSlide = PowerPointSlide.FromSlideFactory(PpOperations.GetCurrentSlide());
            Shape callout = CreateTooltip.GenerateCalloutWithReferenceTriggerShape(currentSlide, selectedShape);
            ConvertToTooltip.AddTriggerAnimation(currentSlide, selectedShape, callout);
            AssertIsSame(testSlideNo, expectedSlideNo);
        }

        private void CreateMultipleCalloutAndCompare(ShapeRange selectedShapes, int testSlideNo, int expectedSlideNo)
        {
            PowerPointSlide currentSlide = PowerPointSlide.FromSlideFactory(PpOperations.GetCurrentSlide());

            foreach (Shape shape in selectedShapes)
            {
                Shape callout = CreateTooltip.GenerateCalloutWithReferenceTriggerShape(currentSlide, shape);
                ConvertToTooltip.AddTriggerAnimation(currentSlide, shape, callout);
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
