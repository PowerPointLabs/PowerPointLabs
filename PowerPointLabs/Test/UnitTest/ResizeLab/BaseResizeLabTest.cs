using System;
using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class BaseResizeLabTest : BaseUnitTest
    {
        private readonly Dictionary<string, string> _originalShapeName = new Dictionary<string, string>();

        protected override string GetTestingSlideName()
        {
            return "ResizeLab\\ResizeLab.pptx";
        }

        protected void InitOriginalShapes(int slideNumber, List<string> shapeNames)
        {
            PowerPoint.ShapeRange shapes = GetShapes(slideNumber, shapeNames);

            _originalShapeName.Clear();
            foreach(PowerPoint.Shape shape in shapes)
            {
                PowerPoint.Shape duplicateShape = shape.Duplicate()[1];
                duplicateShape.Top = shape.Top;
                duplicateShape.Left = shape.Left;
                duplicateShape.Name = Guid.NewGuid().ToString();
                _originalShapeName.Add(duplicateShape.Name, shape.Name);
            }
        }

        protected PowerPoint.ShapeRange GetShapes(int slideNumber, IEnumerable<string> shapeNames)
        {
            PpOperations.SelectSlide(slideNumber);
            return PpOperations.SelectShapes(shapeNames);
        }

        protected void RestoreShapes(int slideNumber, IEnumerable<string> shapeNames)
        {
            try
            {
                List<string> duplicatedShapeNames = new List<string>(_originalShapeName.Keys);
                PowerPoint.ShapeRange executedShapes = GetShapes(slideNumber, shapeNames);
                PowerPoint.ShapeRange shapes = GetShapes(slideNumber, duplicatedShapeNames);
                executedShapes.SafeDelete();

                foreach (PowerPoint.Shape shape in shapes)
                {
                    shape.Name = _originalShapeName[shape.Name];
                }
            }
            catch (Exception)
            {
                // ignored
            }
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
