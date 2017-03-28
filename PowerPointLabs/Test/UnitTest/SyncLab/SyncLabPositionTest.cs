using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabPositionTest : BaseSyncLabTest
    {
        private const int OriginalShapesSlideNo = 4;
        private const string UnrotatedRectangle = "Rectangle 3";
        private const string Oval = "Oval 4";
        private const string RotatedArrow = "Right Arrow 5";
        private const string CopyFromShape = "CopyFrom";

        private Shape _formatShape;

        //Results of Operations
        private const int SyncXPositionSlideNo = 5;
        private const int SyncYPositionSlideNo = 6;
        private const int SyncHeightSlideNo = 7;
        private const int SyncWidthSlideNo = 8;

        [TestInitialize]
        public void TestInitialize()
        {
            _formatShape = GetShape(OriginalShapesSlideNo, CopyFromShape);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncXPosition()
        {
            var newShape = GetShape(OriginalShapesSlideNo, UnrotatedRectangle);
            PositionXFormat.SyncFormat(_formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncXPositionSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncYPosition()
        {
            var newShape = GetShape(OriginalShapesSlideNo, UnrotatedRectangle);
            PositionYFormat.SyncFormat(_formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncYPositionSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncHeight()
        {
            var newShape = GetShape(OriginalShapesSlideNo, Oval);
            PositionHeightFormat.SyncFormat(_formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncHeightSlideNo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncWidth()
        {
            var newShape = GetShape(OriginalShapesSlideNo, RotatedArrow);
            PositionWidthFormat.SyncFormat(_formatShape, newShape);

            CompareSlides(OriginalShapesSlideNo, SyncWidthSlideNo);
        }
    }
}
