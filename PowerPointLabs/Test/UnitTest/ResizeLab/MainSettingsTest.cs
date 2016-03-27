using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ResizeLab;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Test.UnitTest.ResizeLab
{
    [TestClass]
    public class MainSettingsTest : BaseResizeLabTest
    {
        private readonly ResizeLabMain _resizeLab = new ResizeLabMain();
        private List<string> _shapeNames;

        private const string WithAspectRatioShapeNames = "withAspectRatio";
        private const string WithoutAspectRatioShapeNames = "withoutAspectRatio";

        private const int AspectRatioSlideNo = 36;

        [TestInitialize]
        public void TestInitialize()
        {
            switch (TestContext.TestName)
            {
                case "TestLockAspectRatio": case "TestUnlockAspectRatio":
                    _shapeNames = new List<string> { WithAspectRatioShapeNames, WithoutAspectRatioShapeNames };
                    InitOriginalShapes(AspectRatioSlideNo, _shapeNames);
                    break;
            }
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            switch (TestContext.TestName)
            {
                case "TestLockAspectRatio": case "TestUnlockAspectRatio":
                    RestoreShapes(AspectRatioSlideNo, _shapeNames);
                    break;
            }
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestLockAspectRatio()
        {
            var shapes = GetShapes(AspectRatioSlideNo, _shapeNames);

            _resizeLab.ChangeShapesAspectRatio(shapes, true);

            foreach (PowerPoint.Shape shape in shapes)
            {
                Assert.AreEqual(shape.LockAspectRatio, MsoTriState.msoTrue);
            }
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestUnlockAspectRatio()
        {
            var shapes = GetShapes(AspectRatioSlideNo, _shapeNames);

            _resizeLab.ChangeShapesAspectRatio(shapes, false);

            foreach (PowerPoint.Shape shape in shapes)
            {
                Assert.AreEqual(shape.LockAspectRatio, MsoTriState.msoFalse);
            }
        }
    }
}
