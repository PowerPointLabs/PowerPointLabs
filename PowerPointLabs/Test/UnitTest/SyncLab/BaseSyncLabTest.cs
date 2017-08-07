using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Test.Util;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class BaseSyncLabTest : BaseUnitTest
    {
        private readonly Dictionary<string, string> _originalShapeName = new Dictionary<string, string>();

        protected override string GetTestingSlideName()
        {
            return "SyncLab\\SyncLab.pptx";
        }

        protected PowerPoint.Shape GetShape(int slideNumber, string shapeName)
        {
            PpOperations.SelectSlide(slideNumber);
            return PpOperations.SelectShape(shapeName)[1];
        }

        protected PowerPoint.ShapeRange GetShapes(int slideNumber, IEnumerable<string> shapeNames)
        {
            PpOperations.SelectSlide(slideNumber);
            return PpOperations.SelectShapes(shapeNames);
        }

        protected void CompareSlides(int actualShapesSlideNo, int expectedShapesSlideNo)
        {
          
            var actualSlide = PpOperations.SelectSlide(actualShapesSlideNo);
            var expectedSlide = PpOperations.SelectSlide(expectedShapesSlideNo);

            SlideUtil.IsSameLooking(actualSlide, expectedSlide);
        }
    }
}
