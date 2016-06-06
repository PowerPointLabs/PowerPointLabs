using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointLabs.PositionsLab;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Utils;
using System.Diagnostics;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class PositionsLabDuplicateAndRotateTest : BasePositionsLabTest
    {
        private List<string> _shapeNames;

        private const int OriginalShapesSlideNo = 3;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabDuplicateAndRotate.pptx";
        }

        [TestInitialize]
        public void TestInitialize()
        {
            PositionsLabMain.InitPositionsLab();

            _shapeNames = new List<string> { };
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDuplicateAndRoate()
        {
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDuplcateAndRoate2()
        {
        }
    }
}
