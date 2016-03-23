using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using Test.Util;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ResizeLab;
using PowerPointLabs.Utils;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class StretchShrinkTest : BaseUnitTest
    {
        private List<String> _originalShapeNames;
        private ResizeLabMain _resizeLab;
        private Dictionary<string, ShapeProperties> _originalShapesProperties;

        private const string RefShapeName= "ref";
        private const string LeftShapeName = "leftOfRef";
        private const string RightShapeName = "rightOfRef";
        private const string OverShapeName = "overRef";

        private const int OriginalShapesSlideNo = 3;
        private const int TestStretchLeftSlideNo = 4;
        private const int TestStretchRightSlideNo = 5;
        private const int TestStretchTopSlideNo = 6;
        private const int TestStretchBottomSlideNo = 7;

        protected override string GetTestingSlideName()
        {
            return "ResizeLab.pptm";
        }

        [TestInitialize]
        public void TestInitialize()
        {
            PpOperations.SelectSlide(OriginalShapesSlideNo);
            _originalShapeNames = new List<String> {RefShapeName, LeftShapeName,
                RightShapeName, OverShapeName};
            InitOriginalShapes();
            _resizeLab = new ResizeLabMain();
        }

        private void InitOriginalShapes()
        {
            _originalShapesProperties = new Dictionary<string, ShapeProperties>();
            var shapes = GetOriginalShapes();
            foreach (PowerPoint.Shape s in shapes)
            {
                PPShape originalPpShape = new PPShape(s);
                _originalShapesProperties.Add(s.Name, new ShapeProperties(s.Name, originalPpShape.Top, originalPpShape.Left, 
                    originalPpShape.AbsoluteWidth, originalPpShape.AbsoluteHeight, originalPpShape.ShapeRotation));
            }
        }

        private PowerPoint.ShapeRange GetOriginalShapes()
        {
            PpOperations.SelectSlide(OriginalShapesSlideNo);
            return PpOperations.SelectShapes(_originalShapeNames);
        }

        private void ResetOriginalShapes()
        {
            var originalShapes = GetOriginalShapes();
            foreach (PowerPoint.Shape originalShape in originalShapes)
            {
                var originalPpShape = new PPShape(originalShape);
                if (!_originalShapesProperties.ContainsKey(originalPpShape.Name))
                {
                    continue;
                }

                var originalProperty = _originalShapesProperties[originalPpShape.Name];
                originalPpShape.ShapeRotation = originalProperty.ShapeRotation;
                originalPpShape.AbsoluteWidth = originalProperty.AbsoluteWidth;
                originalPpShape.AbsoluteHeight = originalProperty.AbsoluteHeight;
                originalPpShape.Top = originalProperty.Top;
                originalPpShape.Left = originalProperty.Left;

                originalPpShape.ResetNodes();
            }
        }

        private void CheckShapes(PowerPoint.ShapeRange expectedShapes, PowerPoint.ShapeRange actualShapes)
        {
            foreach (PowerPoint.Shape expectedShape in expectedShapes)
            {
                PowerPoint.Shape compareWith = null;

                // Look for the corresponding actual shape
                foreach (PowerPoint.Shape actualShape in actualShapes)
                {
                    if (expectedShape.Name.Equals(actualShape.Name))
                    {
                        compareWith = actualShape;
                        break;
                    }
                    
                }

                if (compareWith == null)
                {
                    Assert.Fail("Unable to find corresponding actual shape");
                }

                SlideUtil.IsSameShape(expectedShape, compareWith);
            }
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchLeft()
        {
            var actualShapes = GetOriginalShapes();
            _resizeLab.StretchLeft(actualShapes);

            PpOperations.SelectSlide(TestStretchLeftSlideNo);
            var expectedResultForStretchLeft = PpOperations.SelectShapes(_originalShapeNames);

            CheckShapes(expectedResultForStretchLeft, actualShapes);
            ResetOriginalShapes();
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchRight()
        {
            var actualShapes = GetOriginalShapes();
            _resizeLab.StretchRight(actualShapes);

            PpOperations.SelectSlide(TestStretchRightSlideNo);
            var expectedResultForStretchRight = PpOperations.SelectShapes(_originalShapeNames);

            CheckShapes(expectedResultForStretchRight, actualShapes);
            ResetOriginalShapes();
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchTop()
        {
            var actualShapes = GetOriginalShapes();
            _resizeLab.StretchTop(actualShapes);

            PpOperations.SelectSlide(TestStretchTopSlideNo);
            var expectedResultForStretchTop = PpOperations.SelectShapes(_originalShapeNames);

            CheckShapes(expectedResultForStretchTop, actualShapes);
            ResetOriginalShapes();
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchBottom()
        {
            var actualShapes = GetOriginalShapes();
            _resizeLab.StretchBottom(actualShapes);

            PpOperations.SelectSlide(TestStretchBottomSlideNo);
            var expectedResultForStretchBottom = PpOperations.SelectShapes(_originalShapeNames);

            CheckShapes(expectedResultForStretchBottom, actualShapes);
            ResetOriginalShapes();
        }
    }
}
