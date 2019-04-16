using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Test.UnitTest.CropLab
{
    [TestClass]
    public abstract class BaseCropLabTest : BaseUnitTest
    {
        private readonly Dictionary<string, string> _originalShapeName = new Dictionary<string, string>();

        protected PowerPoint.ShapeRange GetShapes(int slideNumber, IEnumerable<string> shapeNames)
        {
            PpOperations.SelectSlide(slideNumber);
            return PpOperations.SelectShapes(shapeNames);
        }

        protected void CheckShapes(PowerPoint.ShapeRange expectedShapes, PowerPoint.ShapeRange actualShapes)
        {
            foreach (PowerPoint.Shape actualShape in actualShapes)
            {
                bool isFound = false;

                foreach (PowerPoint.Shape expectedShape in expectedShapes)
                {
                    if (!actualShape.Name.Equals(expectedShape.Name)) continue;
                    isFound = true;
                    SlideUtil.IsSameShape(expectedShape, actualShape);
                    break;
                }

                if (!isFound)
                {
                    Assert.Fail("Unable to find corresponding actual shape");
                }
            }
        }
    }

}
