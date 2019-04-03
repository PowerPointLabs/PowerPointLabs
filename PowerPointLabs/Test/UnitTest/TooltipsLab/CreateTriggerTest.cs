using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.TooltipsLab;
using PowerPointLabs.Models;

using Test.Util;
using Microsoft.Office.Interop.PowerPoint;

namespace Test.UnitTest.TooltipsLab
{
    [TestClass]
    public class CreateTriggerTest : BaseUnitTest
    {
        protected override string GetTestingSlideName()
        {
            return "TooltipsLab\\CreateTrigger.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void CreateTrigger()
        {
            TestCreateTrigger_NormalShape();
            TestCreateTrigger_TriggerShape();
            TestCreateTrigger_CalloutShape();
            TestCreateTrigger_MultipleNormalShapes();
            TestCreateTrigger_MultipleAllTypeShapes();
        }

        private const int CreateTriggerNormalShapeTestSlideNo = 4;
        private const int CreateTriggerNormalShapeExpectedSlideNo = 5;
        private const int CreateTriggerTriggerShapeTestSlideNo = 7;
        private const int CreateTriggerTriggerShapeExpectedSlideNo = 8;
        private const int CreateTriggerCalloutShapeTestSlideNo = 10;
        private const int CreateTriggerCalloutShapeExpectedSlideNo = 11;
        private const int CreateTriggerMultipleNormalShapesTestSlideNo = 13;
        private const int CreateTriggerMultipleNormalShapesExpectedSlideNo = 14;
        private const int CreateTriggerMultipleAllTypeShapesTestSlideNo = 16;
        private const int CreateTriggerMultipleAllTypeShapesExpectedSlideNo = 17;
        private const string NormalShapeName = "normalShape";
        private const string TriggerShapeName = "existingTriggerShape";
        private const string CalloutShapeName = "existingCalloutShape";
        private const string MultipleShapePrefix = "select";

        private void TestCreateTrigger_NormalShape()
        {
            PpOperations.SelectSlide(CreateTriggerNormalShapeTestSlideNo);
            ShapeRange selectedShapeRange = PpOperations.SelectShape(NormalShapeName);
            CreateTriggerAndCompare(selectedShapeRange[1], CreateTriggerNormalShapeTestSlideNo, CreateTriggerNormalShapeExpectedSlideNo);
        }

        private void TestCreateTrigger_TriggerShape()
        {
            PpOperations.SelectSlide(CreateTriggerTriggerShapeTestSlideNo);
            ShapeRange selectedShapeRange = PpOperations.SelectShape(TriggerShapeName);
            CreateTriggerAndCompare(selectedShapeRange[1], CreateTriggerTriggerShapeTestSlideNo, CreateTriggerTriggerShapeExpectedSlideNo);
        }

        private void TestCreateTrigger_CalloutShape()
        {
            PpOperations.SelectSlide(CreateTriggerCalloutShapeTestSlideNo);
            ShapeRange selectedShapeRange = PpOperations.SelectShape(CalloutShapeName);
            CreateTriggerAndCompare(selectedShapeRange[1], CreateTriggerCalloutShapeTestSlideNo, CreateTriggerCalloutShapeExpectedSlideNo);
        }

        private void TestCreateTrigger_MultipleNormalShapes()
        {
            PpOperations.SelectSlide(CreateTriggerMultipleNormalShapesTestSlideNo);
            ShapeRange selectedShapeRange = PpOperations.SelectShapesByPrefix(NormalShapeName);
            CreateMultipleTriggerAndCompare(selectedShapeRange, CreateTriggerMultipleNormalShapesTestSlideNo, CreateTriggerMultipleNormalShapesExpectedSlideNo);
        }

        private void TestCreateTrigger_MultipleAllTypeShapes()
        {
            PpOperations.SelectSlide(CreateTriggerMultipleAllTypeShapesTestSlideNo);
            ShapeRange selectedShapeRange = PpOperations.SelectShapesByPrefix(MultipleShapePrefix);
            CreateMultipleTriggerAndCompare(selectedShapeRange, CreateTriggerMultipleAllTypeShapesTestSlideNo, CreateTriggerMultipleAllTypeShapesExpectedSlideNo);
        }


        private void CreateTriggerAndCompare(Shape selectedShape, int testSlideNo, int expectedSlideNo)
        {
            PowerPointSlide currentSlide = PowerPointSlide.FromSlideFactory(PpOperations.GetCurrentSlide());

            Shape triggerShape = CreateTooltip.GenerateTriggerShapeWithReferenceCallout(currentSlide, selectedShape);
            ConvertToTooltip.AddTriggerAnimation(currentSlide, triggerShape, selectedShape);
            AssertIsSame(testSlideNo, expectedSlideNo);
        }

        private void CreateMultipleTriggerAndCompare(ShapeRange selectedShapes, int testSlideNo, int expectedSlideNo)
        {
            PowerPointSlide currentSlide = PowerPointSlide.FromSlideFactory(PpOperations.GetCurrentSlide());

            foreach (Shape shape in selectedShapes)
            {
                Shape triggerShape = CreateTooltip.GenerateTriggerShapeWithReferenceCallout(currentSlide, shape);
                ConvertToTooltip.AddTriggerAnimation(currentSlide, triggerShape, shape);
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
