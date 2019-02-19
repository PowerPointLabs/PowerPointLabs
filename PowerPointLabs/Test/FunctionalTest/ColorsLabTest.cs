﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

using TestInterface;

using Point = System.Windows.Point;
using Button = System.Windows.Controls.Button;
using System.Threading;

namespace Test.FunctionalTest
{
    [TestClass]
    public class ColorsLabTest : BaseFunctionalTest
    {
        private const int OriginalSlideNo = 3;
        private const int FontColorChangeSlideNo = 4;
        private const int OutlineColorChangeSlideNo = 5;
        private const int FillColorChangeSlideNo = 6;
        private const int BrightnessAndSaturationChangeSlideNo = 7;
        private const int MonochromeColorChangeSlideNo = 8;
        private const int AnalogousAndComplementaryChangeSlideNo = 9;
        private const int TriadicAndTetradicChangeSlideNo = 10;

        private const string TargetShape = "selectMe";
        private const string FontColorShape = "fontColor";
        private const string LineColorShape = "lineColor";
        private const string FillColorShape = "fillColor";

        private List<Color> DefaultTestColors = new List<Color>(new Color[] {
            Color.FromArgb(255, 0, 0),
            Color.FromArgb(0, 255, 0),
            Color.FromArgb(0, 0, 255),
            Color.FromArgb(255, 255, 0),
            Color.FromArgb(255, 0, 255),
            Color.FromArgb(0, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(0, 0, 0),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255)
        });

        private List<Color> AllWhiteColorList = new List<Color>(new Color[]
        {
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255),
            Color.FromArgb(255, 255, 255)
        });

        protected override string GetTestingSlideName()
        {
            return "ColorsLab\\ColorsLab.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_ColorsLabTest()
        {
            // if not maximized, some elements in Colors pane may not be seen
            PpOperations.MaximizeWindow();
            IColorsLabController colorsLab = PplFeatures.ColorsLab;
            colorsLab.OpenPane();

            TestApplyFontColor(colorsLab);
            TestApplyLineColor(colorsLab);
            TestApplyFillColor(colorsLab);

            TestBrightnessAndSaturationSlider(colorsLab);

            TestMonochromeMatchingColors(colorsLab);
            TestAnalogousAndComplementaryColors(colorsLab);
            TestTriadicAndTetradicColors(colorsLab);

            TestFavoriteColors(colorsLab);
        }

        private void TestApplyFontColor(IColorsLabController colorsLab)
        {
            Slide actualSlide = PpOperations.SelectSlide(OriginalSlideNo);

            Shape shape = PpOperations.SelectShape(FontColorShape)[1];
            Point startPt = colorsLab.GetApplyTextButtonLocation();
            Point endPt = GetShapeCenterPoint(shape);

            PpOperations.SelectShape(TargetShape);
            DragAndDrop(startPt, endPt);

            Slide expSlide = PpOperations.SelectSlide(FontColorChangeSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void TestApplyLineColor(IColorsLabController colorsLab)
        {
            Slide actualSlide = PpOperations.SelectSlide(OriginalSlideNo);

            Shape shape = PpOperations.SelectShape(LineColorShape)[1];
            Point startPt = colorsLab.GetApplyLineButtonLocation();
            Point endPt = GetShapeCenterPoint(shape);

            PpOperations.SelectShape(TargetShape);
            DragAndDrop(startPt, endPt);

            Slide expSlide = PpOperations.SelectSlide(OutlineColorChangeSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void TestApplyFillColor(IColorsLabController colorsLab)
        {
            Slide actualSlide = PpOperations.SelectSlide(OriginalSlideNo);

            Shape shape = PpOperations.SelectShape(FillColorShape)[1];
            Point startPt = colorsLab.GetApplyFillButtonLocation();
            Point endPt = GetShapeCenterPoint(shape);

            PpOperations.SelectShape(TargetShape);
            DragAndDrop(startPt, endPt);

            Slide expSlide = PpOperations.SelectSlide(FillColorChangeSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void TestBrightnessAndSaturationSlider(IColorsLabController colorsLab)
        {
            Slide actualSlide = PpOperations.SelectSlide(OriginalSlideNo);

            colorsLab.SlideBrightnessSlider(120);
            colorsLab.SlideSaturationSlider(240);

            Point startPt = colorsLab.GetApplyLineButtonLocation();
            Point endPt = colorsLab.GetMainColorRectangleLocation();

            PpOperations.SelectShape(TargetShape);
            DragAndDrop(startPt, endPt);

            Slide expSlide = PpOperations.SelectSlide(BrightnessAndSaturationChangeSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void TestMonochromeMatchingColors(IColorsLabController colorsLab)
        {
            Slide actualSlide = PpOperations.SelectSlide(OriginalSlideNo);

            // Apply monochromeRectOne as Line
            colorsLab.ClickMonochromeRect(1);
            Point startPt = colorsLab.GetApplyLineButtonLocation();
            Point endPt = colorsLab.GetMainColorRectangleLocation();
            PpOperations.SelectShape(TargetShape);
            DragAndDrop(startPt, endPt);

            // Apply monochromeRectTwo as Text
            colorsLab.ClickMonochromeRect(2);
            startPt = colorsLab.GetApplyTextButtonLocation();
            endPt = colorsLab.GetMainColorRectangleLocation();
            PpOperations.SelectShape(TargetShape);
            DragAndDrop(startPt, endPt);

            // Apply monochromeRectSix as Fill
            colorsLab.ClickMonochromeRect(6);
            startPt = colorsLab.GetApplyFillButtonLocation();
            endPt = colorsLab.GetMainColorRectangleLocation();
            PpOperations.SelectShape(TargetShape);
            DragAndDrop(startPt, endPt);

            Slide expSlide = PpOperations.SelectSlide(MonochromeColorChangeSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }


        private void TestAnalogousAndComplementaryColors(IColorsLabController colorsLab)
        {
            Slide actualSlide = PpOperations.SelectSlide(OriginalSlideNo);

            // Eyedrop the font color shape
            Shape shape = PpOperations.SelectShape(FontColorShape)[1];
            Point startPt = colorsLab.GetEyeDropperButtonLocation();
            Point endPt = GetShapeCenterPoint(shape);
            DragAndDrop(startPt, endPt);

            // Apply analagousRectOne as Text
            colorsLab.ClickAnalogousRect(1);
            startPt = colorsLab.GetApplyTextButtonLocation();
            endPt = colorsLab.GetMainColorRectangleLocation();
            PpOperations.SelectShape(TargetShape);
            DragAndDrop(startPt, endPt);

            // Apply complementaryRectThree as Line
            colorsLab.ClickComplementaryRect(3);
            startPt = colorsLab.GetApplyLineButtonLocation();
            endPt = colorsLab.GetMainColorRectangleLocation();
            PpOperations.SelectShape(TargetShape);
            DragAndDrop(startPt, endPt);

            Slide expSlide = PpOperations.SelectSlide(AnalogousAndComplementaryChangeSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void TestTriadicAndTetradicColors(IColorsLabController colorsLab)
        {
            Slide actualSlide = PpOperations.SelectSlide(OriginalSlideNo);

            // Eyedrop the outline color shape
            Shape shape = PpOperations.SelectShape(LineColorShape)[1];
            Point startPt = colorsLab.GetEyeDropperButtonLocation();
            Point endPt = GetShapeCenterPoint(shape);
            DragAndDrop(startPt, endPt);

            // Apply triadicRectThree as Fill
            colorsLab.ClickTriadicRect(3);
            startPt = colorsLab.GetApplyFillButtonLocation();
            endPt = colorsLab.GetMainColorRectangleLocation();
            PpOperations.SelectShape(TargetShape);
            DragAndDrop(startPt, endPt);

            // Apply tetradicRectFour as Line
            colorsLab.ClickTetradicRect(4);
            startPt = colorsLab.GetApplyLineButtonLocation();
            endPt = colorsLab.GetMainColorRectangleLocation();
            PpOperations.SelectShape(TargetShape);
            DragAndDrop(startPt, endPt);

            Slide expSlide = PpOperations.SelectSlide(TriadicAndTetradicChangeSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }


        private void TestFavoriteColors(IColorsLabController colorsLab)
        {
            // Clear the favorite colors panel
            colorsLab.ClearFavoriteColors();
            List<Color> currentFavoritePanel = colorsLab.GetCurrentFavoritePanel();
            AssertEqual(AllWhiteColorList, currentFavoritePanel);

            // Load the test case
            colorsLab.LoadFavoriteColors(
                PathUtil.GetDocTestPresentationPath("ColorsLab\\FavoriteColorsTest.thm"));
            currentFavoritePanel = colorsLab.GetCurrentFavoritePanel();
            AssertEqual(DefaultTestColors, currentFavoritePanel);

            // Clear panel
            colorsLab.ClearFavoriteColors();
            currentFavoritePanel = colorsLab.GetCurrentFavoritePanel();
            AssertEqual(AllWhiteColorList, currentFavoritePanel);
        }

        # region Helper methods
        // mouse drag & drop from Control to Shape to apply color
        private void ApplyColor(System.Windows.Controls.Control from, Shape to)
        {
            Point startPt = from.PointToScreen(new Point(from.Width/2, from.Height/2));
            Point endPt = new Point(
                PpOperations.PointsToScreenPixelsX(to.Left + to.Width/2),
                PpOperations.PointsToScreenPixelsY(to.Top + to.Height/2));
            DragAndDrop(startPt, endPt);
        }

        private Point GetShapeCenterPoint(Shape shape)
        {
            return new Point(
                PpOperations.PointsToScreenPixelsX(shape.Left + shape.Width / 2),
                PpOperations.PointsToScreenPixelsY(shape.Top + shape.Height / 2));
        }

        private void DragAndDrop(Point startPt, Point endPt)
        {
            MouseUtil.SendMouseDown((int)startPt.X, (int)startPt.Y);
            MouseUtil.SendMouseUp((int)endPt.X, (int)endPt.Y);
        }



        private static void AssertEqual(List<Color> exp, List<Color> actual)
        {
            for (int i = 0; i < 13; i++)
            {
                AssertEqual(exp[i], actual[i]);
            }
        }

        private static void AssertEqual(Color exp, Color actual)
        {
            // dont assert Alpha
            Assert.IsTrue(IsAlmostSame(exp.R, actual.R), "diff color R, exp {0}, actual {1}", exp.R, actual.R);
            Assert.IsTrue(IsAlmostSame(exp.G, actual.G),"diff color G, exp {0}, actual {1}", exp.G, actual.G);
            Assert.IsTrue(IsAlmostSame(exp.B, actual.B), "diff color B, exp {0}, actual {1}", exp.B, actual.B);
        }

        private static bool IsAlmostSame(byte a, byte b)
        {
            return Math.Abs(a - b) < 3;
        }
        # endregion
    }
}
