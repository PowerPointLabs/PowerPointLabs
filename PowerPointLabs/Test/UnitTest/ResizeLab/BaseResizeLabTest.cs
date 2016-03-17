using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Test.Util;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ResizeLab;
using PowerPointLabs.Utils;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class BaseResizeLabTest : BaseUnitTest
    {
        protected override string GetTestingSlideName()
        {
            return "ResizeLab.pptm";
        }

        protected void InitOriginalShapes(int slideNumber, List<string> shapeNames,
            Dictionary<string, ShapeProperties> shapeProperties)
        {
            var shapes = GetShapes(slideNumber, shapeNames);
            foreach (PowerPoint.Shape shape in shapes)
            {
                var originalPpShape = new PPShape(shape);
                shapeProperties.Add(shape.Name,
                    new ShapeProperties(shape.Name, originalPpShape.Top, originalPpShape.Left,
                        originalPpShape.AbsoluteWidth, originalPpShape.AbsoluteHeight));
            }
        }

        protected PowerPoint.ShapeRange GetShapes(int slideNumber, IEnumerable<string> shapeNames)
        {
            PpOperations.SelectSlide(slideNumber);
            return PpOperations.SelectShapes(shapeNames);
        }

        protected void RestoreShapes(PowerPoint.ShapeRange shapes, IDictionary<string, ShapeProperties> shapeProperties)
        {
            foreach (PowerPoint.Shape shape in shapes)
            {
                var ppShape = new PPShape(shape);
                if (!shapeProperties.ContainsKey(ppShape.Name))
                {
                    continue;
                }

                var originalProperty = shapeProperties[ppShape.Name];
                ppShape.Top = originalProperty.Top;
                ppShape.Left = originalProperty.Left;
                ppShape.AbsoluteWidth = originalProperty.AbsoluteWidth;
                ppShape.AbsoluteHeight = originalProperty.AbsoluteHeight;
            }
        }

        protected void CheckShapes(PowerPoint.ShapeRange expectedShapes, PowerPoint.ShapeRange actualShapes)
        {
            foreach (PowerPoint.Shape actualShape in actualShapes)
            {
                var isFound = false;

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
