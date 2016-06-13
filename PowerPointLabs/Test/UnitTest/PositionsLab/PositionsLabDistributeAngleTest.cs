using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using PowerPointLabs.PositionsLab;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Utils;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class PositionsLabDistributeAngleTest : BasePositionsLabTest
    {
        private List<string> _shapeNames;

        private const int OriginalShapesSlideNo = 3;
        private const string Circle = "Oval 10";
        private const string Pie = "Pie 2";
        private const string Arrow = "Circular Arrow 9";
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
        private const int DistributeAngleAtSecondWithEdgesSlide = 5;
        private const int DistributeAngleAtSecondWithCenterSlide = 6;

        private const int DistributeAngleWithinSecondWithEdgesSlide = 8;
        private const int DistributeAngleWithinSecondWithCenterSlide = 9;

        private const int DistributeAngleWithinSecondAndThirdWithEdgesSlide = 11;
        private const int DistributeAngleWithinSecondAndThirdWithCenterSlide = 12;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabDistributeAngle.pptx";
        }

        [TestInitialize]
        public void TestInitialize()
        {
            PositionsLabMain.InitPositionsLab();

            _shapeNames = new List<string> { Circle, Pie, Arrow, WhiteChevron, RedChevron, OrangeChevron, YellowChevron, GreenChevron,
                BlueChevron, PurpleChevron, BlackChevron, Picture};
            InitOriginalShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestCleanup]
        public void TestCleanUp()
        {
            RestoreShapes(OriginalShapesSlideNo, _shapeNames);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleAtSecondWithEdges()
        {
            PositionsLabMain.DistributeReferAtSecondShape();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { Circle, Pie, WhiteChevron, RedChevron, OrangeChevron, YellowChevron, GreenChevron, BlueChevron,
                PurpleChevron, BlackChevron, Picture };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeAngle(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleAtSecondWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleAtSecondWithCenter()
        {
            PositionsLabMain.DistributeReferAtSecondShape();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { Circle, Pie, WhiteChevron, RedChevron, OrangeChevron, YellowChevron, GreenChevron, BlueChevron,
                PurpleChevron, BlackChevron, Picture };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeAngle(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleAtSecondWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleWithinSecondWithEdges()
        {
            PositionsLabMain.DistributeReferToSecondShape();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { Circle, Arrow, WhiteChevron, RedChevron, OrangeChevron, YellowChevron, GreenChevron, BlueChevron,
                PurpleChevron, BlackChevron, Picture };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeAngle(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleWithinSecondWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleWithinSecondWithCenter()
        {
            PositionsLabMain.DistributeReferToSecondShape();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { Circle, Arrow, WhiteChevron, RedChevron, OrangeChevron, YellowChevron, GreenChevron, BlueChevron,
                PurpleChevron, BlackChevron, Picture };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeAngle(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleWithinSecondWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleWithinSecondAndThirdWithEdges()
        {
            PositionsLabMain.DistributeReferToSecondThirdShape();
            PositionsLabMain.DistributeSpaceByBoundaries();
            _shapeNames = new List<string> { Circle, Pie, PurpleChevron, WhiteChevron, RedChevron, OrangeChevron, YellowChevron, GreenChevron,
                BlueChevron, BlackChevron, Picture };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeAngle(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleWithinSecondAndThirdWithEdgesSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleWithinSecondAndThirdWithCenter()
        {
            PositionsLabMain.DistributeReferToSecondThirdShape();
            PositionsLabMain.DistributeSpaceByCenter();
            _shapeNames = new List<string> { Circle, Pie, PurpleChevron, WhiteChevron, RedChevron, OrangeChevron, YellowChevron, GreenChevron,
                BlueChevron, BlackChevron, Picture };
            var actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeAngle(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleWithinSecondAndThirdWithCenterSlide);
            var expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
