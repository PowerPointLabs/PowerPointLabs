using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.Models;
using PowerPointLabs.TooltipsLab;
using PowerPointLabs.Utils;
using Test.Util;

namespace Test.UnitTest.TooltipsLab
{
    [TestClass]
    public class PPShapeTest : BaseUnitTest
    {
        private const int UnrotatedNoTextTestSlideNo = 1;
        private const int UnrotatedNoTextExpectedSlideNo = 2;
        private const int RotatedNoTextTestSlideNo = 3;
        private const int RotatedNoTextExpectedSlideNo = 4;
        private const int UnrotatedTextTestSlideNo = 5;
        private const int UnrotatedTextExpectedSlideNo = 6;
        private const int RotatedTextTestSlideNo = 7;
        private const int RotatedTextExpectedSlideNo = 8;
        private const int UnrotatedVerticalTextTestSlideNo = 9;
        private const int UnrotatedVerticalTextExpectedSlideNo = 10;
        private const int RotatedVerticalTextTestSlideNo = 11;
        private const int RotatedVerticalTextExpectedSlideNo = 12;
        private const int UnrotatedRotatedTextTestSlideNo = 13;
        private const int UnrotatedRotatedTextExpectedSlideNo = 14;
        private const int RotatedRotatedTextTestSlideNo = 15;
        private const int RotatedRotatedTextExpectedSlideNo = 16;

        private const string ShapeName = "Shape";

        protected override string GetTestingSlideName()
        {
            return "PPShapes.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void CreatePPShape()
        {
            TestCreatePPShape(UnrotatedNoTextTestSlideNo, UnrotatedNoTextExpectedSlideNo);
            TestCreatePPShape(RotatedNoTextTestSlideNo, RotatedNoTextExpectedSlideNo);
            TestCreatePPShape(UnrotatedTextTestSlideNo, UnrotatedTextExpectedSlideNo);
            TestCreatePPShape(RotatedTextTestSlideNo, RotatedTextExpectedSlideNo);
            TestCreatePPShape(UnrotatedVerticalTextTestSlideNo, UnrotatedVerticalTextExpectedSlideNo);
            TestCreatePPShape(RotatedVerticalTextTestSlideNo, RotatedVerticalTextExpectedSlideNo);
            TestCreatePPShape(UnrotatedRotatedTextTestSlideNo, UnrotatedRotatedTextExpectedSlideNo);
            TestCreatePPShape(RotatedRotatedTextTestSlideNo, RotatedRotatedTextExpectedSlideNo);
        }

        private void TestCreatePPShape(int testSlideNo, int expectedSlideNo)
        {
            PpOperations.SelectSlide(UnrotatedNoTextTestSlideNo);
            Shape selectedShape = PpOperations.SelectShape(ShapeName)[1];
            PPShape ppShape = new PPShape(selectedShape);
            AssertIsSame(testSlideNo, expectedSlideNo);
        }

        private void AssertIsSame(int actualSlideIndex, int expectedSlideIndex)
        {
            Slide actualSlide = PpOperations.SelectSlide(actualSlideIndex);
            Slide expectedSlide = PpOperations.SelectSlide(expectedSlideIndex);
            Util.SlideUtil.IsSameLooking(expectedSlide, actualSlide, 0.99);
        }
    }
}
