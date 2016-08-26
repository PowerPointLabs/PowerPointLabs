using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointLabs.PositionsLab;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Utils;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class PositionsLabDistributeRadialTest : BasePositionsLabTest
    {
        private List<string> _shapeNames;

        private const int OriginalShapesSlideNo = 3;
        private const string Circle = "Oval 10";
        private const string Pie = "Pie 2";
        private const string WhiteChevron = "Chevron 1";
        private const string RedChevron = "Chevron 20";
        private const string OrangeChevron = "Chevron 19";
        private const string YellowChevron = "Chevron 18";
        private const string GreenChevron = "Chevron 17";
        private const string BlueChevron = "Chevron 16";
        private const string PurpleChevron = "Chevron 15";
        private const string BlackChevron = "Chevron 14";
        private const string Picture = "Picture 4";

        //Results of Operations
        private const int DistributeAngleAtSecondWithEdgesFixedShapeOrientationSlide = 5;
        private const int DistributeAngleAtSecondWithEdgesDynamicShapeOrientationSlide = 6;
        private const int DistributeAngleAtSecondWithCenterFixedShapeOrientationSlide = 7;
        private const int DistributeAngleAtSecondWithCenterDynamicShapeOrientationSlide = 8;

        private const int DistributeAngleWithinSecondAndThirdWithEdgesFixedShapeOrientationSlide = 10;
        private const int DistributeAngleWithinSecondAndThirdWithEdgesDynamicShapeOrientationSlide = 11;
        private const int DistributeAngleWithinSecondAndThirdWithCenterFixedShapeOrientationSlide = 12;
        private const int DistributeAngleWithinSecondAndThirdWithCenterDynamicShapeOrientationSlide = 13;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabDistributeRadial.pptx";
        }

        [TestInitialize]
        public void TestInitialize()
        {
            PositionsLabMain.InitPositionsLab();

            _shapeNames = new List<string> { Circle, Pie, PurpleChevron, BlackChevron, BlueChevron, GreenChevron, OrangeChevron,
                Picture, RedChevron, WhiteChevron, YellowChevron};
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleAtSecondWithEdgesFixedShapeOrientation()
        {
            PositionsLabMain.DistributeReferAtSecondShape();
            PositionsLabMain.DistributeSpaceByBoundaries();
            PositionsLabMain.DistributeShapeOrientationToFixed();
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleAtSecondWithEdgesFixedShapeOrientationSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleAtSecondWithEdgesDynamicShapeOrientation()
        {
            PositionsLabMain.DistributeReferAtSecondShape();
            PositionsLabMain.DistributeSpaceByBoundaries();
            PositionsLabMain.DistributeShapeOrientationToDynamic();
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleAtSecondWithEdgesDynamicShapeOrientationSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleAtSecondWithCenterFixedShapeOrientation()
        {
            PositionsLabMain.DistributeReferAtSecondShape();
            PositionsLabMain.DistributeSpaceByCenter();
            PositionsLabMain.DistributeShapeOrientationToFixed();
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleAtSecondWithCenterFixedShapeOrientationSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleAtSecondWithCenterDynamicShapeOrientation()
        {
            PositionsLabMain.DistributeReferAtSecondShape();
            PositionsLabMain.DistributeSpaceByCenter();
            PositionsLabMain.DistributeShapeOrientationToDynamic();
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleAtSecondWithCenterDynamicShapeOrientationSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleWithinSecondAndThirdWithEdgesFixedShapeOrientation()
        {
            PositionsLabMain.DistributeReferToSecondThirdShape();
            PositionsLabMain.DistributeSpaceByBoundaries();
            PositionsLabMain.DistributeShapeOrientationToFixed();
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleWithinSecondAndThirdWithEdgesFixedShapeOrientationSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleWithinSecondAndThirdWithEdgesDynamicShapeOrientation()
        {
            PositionsLabMain.DistributeReferToSecondThirdShape();
            PositionsLabMain.DistributeSpaceByBoundaries();
            PositionsLabMain.DistributeShapeOrientationToDynamic();
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleWithinSecondAndThirdWithEdgesDynamicShapeOrientationSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleWithinSecondAndThirdWithCenterFixedShapeOrientation()
        {
            PositionsLabMain.DistributeReferToSecondThirdShape();
            PositionsLabMain.DistributeSpaceByCenter();
            PositionsLabMain.DistributeShapeOrientationToFixed();
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleWithinSecondAndThirdWithCenterFixedShapeOrientationSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleWithinSecondAndThirdWithCenterDynamicShapeOrientation()
        {
            PositionsLabMain.DistributeReferToSecondThirdShape();
            PositionsLabMain.DistributeSpaceByCenter();
            PositionsLabMain.DistributeShapeOrientationToDynamic();
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleWithinSecondAndThirdWithCenterDynamicShapeOrientationSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
