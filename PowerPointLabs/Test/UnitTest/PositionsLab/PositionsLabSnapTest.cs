using System;
using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PositionsLab;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace Test.UnitTest.PositionsLab
{
    [TestClass]
    public class PositionsLabSnapTest : BasePositionsLabTest
    {
        private List<string> _shapeNames;

        private const int OriginalShapes0DegreesSlideNo = 3;
        private const int OriginalShapes315DegreesSlideNo = 4;
        private const int OriginalShapes225DegreesSlideNo = 5;
        private const int OriginalShapes135DegreesSlideNo = 6;
        private const int OriginalShapes45DegreesSlideNo = 7;
        private const string Rectangle3 = "Rectangle 3";
        private const string Oval4 = "Oval 4";
        private const string Rectangle6 = "Rectangle 6";
        private const string Picture3 = "Picture 3";
        private const string Picture4 = "Picture 4";
        private const string Picture2 = "Picture 2";
        private const string LeftArrow2 = "Left Arrow 2";
        private const string UpArrow10 = "Up Arrow 10";
        private const string DownArrow11 = "Down Arrow 11";
        private const string LeftRightArrow12 = "Left-Right Arrow 12";
        private const string UpDownArrow13 = "Up-Down Arrow 13";
        private const string QuadArrow14 = "Quad Arrow 14";
        private const string LeftRightUpArrow15 = "Left-Right-Up Arrow 15";
        private const string BentArrow16 = "Bent Arrow 16";
        private const string UTurnArrow17 = "U-Turn Arrow 17";
        private const string LeftUpArrow18 = "Left-Up Arrow 18";
        private const string BentUpArrow19 = "Bent-Up Arrow 19";
        private const string CurvedRightArrow20 = "Curved Right Arrow 20";
        private const string CurvedLeftArrow21 = "Curved Left Arrow 21";
        private const string CurvedDownArrow22 = "Curved Down Arrow 22";
        private const string CurvedUpArrow23 = "Curved Up Arrow 23";
        private const string StripedRightArrow24 = "Striped Right Arrow 24";
        private const string NotchedRightArrow25 = "Notched Right Arrow 25";
        private const string Pentagon26 = "Pentagon 26";
        private const string Chevron27 = "Chevron 27";
        private const string RightArrowCallout28 = "Right Arrow Callout 28";
        private const string DownArrowCallout29 = "Down Arrow Callout 29";
        private const string LeftArrowCallout30 = "Left Arrow Callout 30";
        private const string UpArrowCallout31 = "Up Arrow Callout 31";
        private const string LeftRightArrowCallout32 = "Left-Right Arrow Callout 32";
        private const string QuadArrowCallout33 = "Quad Arrow Callout 33";
        private const string CircularArrow34 = "Circular Arrow 34";
        private const string RightArrow1 = "Right Arrow 1";

        //Results of Operations
        private const int SnapHorizontal0DegreesSlideNo = 9;
        private const int SnapHorizontal315DegreesSlideNo = 10;
        private const int SnapHorizontal225DegreesSlideNo = 11;
        private const int SnapHorizontal135DegreesSlideNo = 12;
        private const int SnapHorizontal45DegreesSlideNo = 13;

        private const int SnapVertical0DegreesSlideNo = 15;
        private const int SnapVertical315DegreesSlideNo = 16;
        private const int SnapVertical225DegreesSlideNo = 17;
        private const int SnapVertical135DegreesSlideNo = 18;
        private const int SnapVertical45DegreesSlideNo = 19;

        private const int SnapAway1SlideNo = 21;
        private const int SnapAway2SlideNo = 22;
        private const int SnapAway3SlideNo = 23;
        private const int SnapAway4SlideNo = 24;

        protected override string GetTestingSlideName()
        {
            return "PositionsLab\\PositionsLabSnap.pptx";
        }

        [TestInitialize]
        public void TestInitialize()
        {
            PositionsLabMain.InitPositionsLab();
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSnapHorizontal0Degrees()
        {
            _shapeNames = new List<string> { Rectangle3, Oval4, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapes0DegreesSlideNo, _shapeNames);

            Action<IList<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapHorizontal(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(SnapHorizontal0DegreesSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSnapHorizontal315Degrees()
        {
            _shapeNames = new List<string> { Rectangle3, Oval4, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapes315DegreesSlideNo, _shapeNames);

            Action<IList<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapHorizontal(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(SnapHorizontal315DegreesSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSnapHorizontal225Degrees()
        {
            _shapeNames = new List<string> { Rectangle3, Oval4, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapes225DegreesSlideNo, _shapeNames);

            Action<IList<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapHorizontal(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(SnapHorizontal225DegreesSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSnapHorizontal135Degrees()
        {
            _shapeNames = new List<string> { Rectangle3, Oval4, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapes135DegreesSlideNo, _shapeNames);

            Action<IList<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapHorizontal(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(SnapHorizontal135DegreesSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSnapHorizontal45Degrees()
        {
            _shapeNames = new List<string> { Rectangle3, Oval4, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapes45DegreesSlideNo, _shapeNames);

            Action<IList<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapHorizontal(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(SnapHorizontal45DegreesSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSnapVertical0Degrees()
        {
            _shapeNames = new List<string> { Rectangle3, Oval4, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapes0DegreesSlideNo, _shapeNames);

            Action<IList<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapVertical(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(SnapVertical0DegreesSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSnapVertical315Degrees()
        {
            _shapeNames = new List<string> { Rectangle3, Oval4, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapes315DegreesSlideNo, _shapeNames);

            Action<IList<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapVertical(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(SnapVertical315DegreesSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSnapVertical225Degrees()
        {
            _shapeNames = new List<string> { Rectangle3, Oval4, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapes225DegreesSlideNo, _shapeNames);

            Action<IList<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapVertical(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(SnapVertical225DegreesSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSnapVertical135Degrees()
        {
            _shapeNames = new List<string> { Rectangle3, Oval4, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapes135DegreesSlideNo, _shapeNames);

            Action<IList<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapVertical(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(SnapVertical135DegreesSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSnapVertical45Degrees()
        {
            _shapeNames = new List<string> { Rectangle3, Oval4, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapes45DegreesSlideNo, _shapeNames);

            Action<IList<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapVertical(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(SnapVertical45DegreesSlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSnapAway1()
        {
            _shapeNames = new List<string> { Oval4, Rectangle3, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapes0DegreesSlideNo, _shapeNames);

            Action<IList<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapAway(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(SnapAway1SlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSnapAway2()
        {
            _shapeNames = new List<string> { Oval4, Rectangle3, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapes0DegreesSlideNo, _shapeNames);

            Action<IList<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapAway(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(SnapAway2SlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSnapAway3()
        {
            _shapeNames = new List<string> { Oval4, Rectangle3, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapes0DegreesSlideNo, _shapeNames);

            Action<IList<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapAway(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);
            ExecutePositionsAction(positionsAction, actualShapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(SnapAway3SlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSnapAway4()
        {
            _shapeNames = new List<string> { Oval4, Rectangle3, Rectangle6, Picture3, Picture4, Picture2, LeftArrow2, UpArrow10, DownArrow11,
                                            LeftRightArrow12, UpDownArrow13, QuadArrow14, LeftRightUpArrow15, BentArrow16, UTurnArrow17, LeftUpArrow18, BentUpArrow19,
                                            CurvedRightArrow20, CurvedLeftArrow21, CurvedDownArrow22, CurvedUpArrow23, StripedRightArrow24, NotchedRightArrow25, Pentagon26,
                                            Chevron27, RightArrowCallout28, DownArrowCallout29, LeftArrowCallout30, UpArrowCallout31, LeftRightArrowCallout32, UpArrowCallout31,
                                            LeftRightArrowCallout32, QuadArrowCallout33, CircularArrow34, RightArrow1 };
            PowerPoint.ShapeRange actualShapes = GetShapes(OriginalShapes0DegreesSlideNo, _shapeNames);

            Action<IList<Shape>> positionsAction = (shapes) => PositionsLabMain.SnapAway(shapes);
            ExecutePositionsAction(positionsAction, actualShapes);
            ExecutePositionsAction(positionsAction, actualShapes);
            ExecutePositionsAction(positionsAction, actualShapes);
            ExecutePositionsAction(positionsAction, actualShapes);

            PpOperations.SelectSlide(SnapAway4SlideNo);
            PowerPoint.ShapeRange expectedShapes = PpOperations.SelectShapes(_shapeNames);

            CheckShapes(expectedShapes, actualShapes);
        }
    }
}
