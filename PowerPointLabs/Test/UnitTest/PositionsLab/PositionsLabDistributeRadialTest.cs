using System;
using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PositionsLab;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

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
            PositionsLabSettings.DistributeRadialReference = PositionsLabSettings.DistributeRadialReferenceObject.AtSecondShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary;
            PositionsLabSettings.DistributeShapeOrientation = PositionsLabSettings.RadialShapeOrientationObject.Fixed;

            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleAtSecondWithEdgesFixedShapeOrientationSlide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleAtSecondWithEdgesDynamicShapeOrientation()
        {
            PositionsLabSettings.DistributeRadialReference = PositionsLabSettings.DistributeRadialReferenceObject.AtSecondShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary;
            PositionsLabSettings.DistributeShapeOrientation = PositionsLabSettings.RadialShapeOrientationObject.Dynamic;

            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleAtSecondWithEdgesDynamicShapeOrientationSlide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleAtSecondWithCenterFixedShapeOrientation()
        {
            PositionsLabSettings.DistributeRadialReference = PositionsLabSettings.DistributeRadialReferenceObject.AtSecondShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            PositionsLabSettings.DistributeShapeOrientation = PositionsLabSettings.RadialShapeOrientationObject.Fixed;

            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleAtSecondWithCenterFixedShapeOrientationSlide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleAtSecondWithCenterDynamicShapeOrientation()
        {
            PositionsLabSettings.DistributeRadialReference = PositionsLabSettings.DistributeRadialReferenceObject.AtSecondShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            PositionsLabSettings.DistributeShapeOrientation = PositionsLabSettings.RadialShapeOrientationObject.Dynamic;

            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleAtSecondWithCenterDynamicShapeOrientationSlide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleWithinSecondAndThirdWithEdgesFixedShapeOrientation()
        {
            PositionsLabSettings.DistributeRadialReference = PositionsLabSettings.DistributeRadialReferenceObject.SecondThirdShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary;
            PositionsLabSettings.DistributeShapeOrientation = PositionsLabSettings.RadialShapeOrientationObject.Fixed;

            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleWithinSecondAndThirdWithEdgesFixedShapeOrientationSlide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleWithinSecondAndThirdWithEdgesDynamicShapeOrientation()
        {
            PositionsLabSettings.DistributeRadialReference = PositionsLabSettings.DistributeRadialReferenceObject.SecondThirdShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary;
            PositionsLabSettings.DistributeShapeOrientation = PositionsLabSettings.RadialShapeOrientationObject.Dynamic;

            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleWithinSecondAndThirdWithEdgesDynamicShapeOrientationSlide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleWithinSecondAndThirdWithCenterFixedShapeOrientation()
        {
            PositionsLabSettings.DistributeRadialReference = PositionsLabSettings.DistributeRadialReferenceObject.SecondThirdShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            PositionsLabSettings.DistributeShapeOrientation = PositionsLabSettings.RadialShapeOrientationObject.Fixed;

            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleWithinSecondAndThirdWithCenterFixedShapeOrientationSlide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDistributeAngleWithinSecondAndThirdWithCenterDynamicShapeOrientation()
        {
            PositionsLabSettings.DistributeRadialReference = PositionsLabSettings.DistributeRadialReferenceObject.SecondThirdShape;
            PositionsLabSettings.DistributeSpaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            PositionsLabSettings.DistributeShapeOrientation = PositionsLabSettings.RadialShapeOrientationObject.Dynamic;

            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapesSlideNo, _shapeNames);

            Action<PowerPoint.ShapeRange> positionsAction = (shapes) => PositionsLabMain.DistributeRadial(shapes);
            ExecutePositionsAction(positionsAction, actualShapes, isConvertPPShape: false);

            PpOperations.SelectSlide(DistributeAngleWithinSecondAndThirdWithCenterDynamicShapeOrientationSlide);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
