using System.Collections.Generic;

using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.ResizeLab;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class StretchShrinkTest : BaseResizeLabTest
    {
        private readonly ResizeLabMain _resizeLab = new ResizeLabMain();
        private List<string> _shapeNames;

        private const string RefShapeName= "ref";
        private const string LeftShapeName = "leftOfRef";
        private const string RightShapeName = "rightOfRef";
        private const string OverShapeName = "overRef";

        [TestInitialize]
        public void TestInitialize()
        {
            _shapeNames = new List<string> {RefShapeName, LeftShapeName, RightShapeName, OverShapeName};
            InitOriginalShapes(SlideNo.StretchOrigin, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(SlideNo.StretchOrigin, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchLeftWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.StretchLeft, _shapeNames);

            _resizeLab.StretchLeft(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchLeftWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.StretchLeftAspectRatio, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoTrue;
            _resizeLab.StretchLeft(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchLeftOuterMost()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.StretchLeftOuterMost, _shapeNames);

            _resizeLab.ReferenceType = ResizeLabMain.StretchRefType.Outermost;
            _resizeLab.StretchLeft(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchRightWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.StretchRight, _shapeNames);

            _resizeLab.StretchRight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchRightWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.StretchRightAspectRatio, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoTrue;
            _resizeLab.StretchRight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchRightOuterMost()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.StretchRightOuterMost, _shapeNames);

            _resizeLab.ReferenceType = ResizeLabMain.StretchRefType.Outermost;
            _resizeLab.StretchRight(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchTopWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.StretchTop, _shapeNames);

            _resizeLab.StretchTop(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchTopWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.StretchTopAspectRatio, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoTrue;
            _resizeLab.StretchTop(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchTopOuterMost()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.StretchTopOuterMost, _shapeNames);

            _resizeLab.ReferenceType = ResizeLabMain.StretchRefType.Outermost;
            _resizeLab.StretchTop(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchBottomWithoutAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.StretchBottom, _shapeNames);

            _resizeLab.StretchBottom(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchBottomWithAspectRatio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.StretchBottomAspectRatio, _shapeNames);

            actualShapes.LockAspectRatio = MsoTriState.msoTrue;
            _resizeLab.StretchBottom(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStretchBottomOuterMost()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange actualShapes = GetShapes(SlideNo.StretchOrigin, _shapeNames);
            Microsoft.Office.Interop.PowerPoint.ShapeRange expectedShapes = GetShapes(SlideNo.StretchBottomOuterMost, _shapeNames);

            _resizeLab.ReferenceType = ResizeLabMain.StretchRefType.Outermost;
            _resizeLab.StretchBottom(actualShapes);
            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
