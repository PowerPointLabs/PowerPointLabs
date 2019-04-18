using System;
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
            Color.FromArgb(255, 255, 255)
        });

        private List<Color> RecentColorsAfterFT = new List<Color>(new Color[]
        {
            Color.FromArgb(98, 235, 187),
            Color.FromArgb(118, 98, 235),
            Color.FromArgb(98, 215, 235),
            Color.FromArgb(235, 98, 146),
            Color.FromArgb(0, 53, 153),
            Color.FromArgb(102, 155, 255),
            Color.FromArgb(153, 189, 255),
            Color.FromArgb(0, 89, 255),
            Color.FromArgb(98, 148, 235),
            Color.FromArgb(98, 235, 118),
            Color.FromArgb(235, 118, 98),
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

            // Clear the recent colors panel before FT begins
            colorsLab.ClearRecentColors();

            TestApplyFontColor(colorsLab);
            TestApplyLineColor(colorsLab);
            TestApplyFillColor(colorsLab);

            TestBrightnessAndSaturationSlider(colorsLab);

            TestMonochromeMatchingColors(colorsLab);
            TestAnalogousAndComplementaryColors(colorsLab);
            TestTriadicAndTetradicColors(colorsLab);

            TestFavoriteColors(colorsLab);
            TestRecentColors(colorsLab);
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

            // Apply tetradicRectThree as Line
            colorsLab.ClickTetradicRect(3);
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
            IList<Color> currentFavoritePanel = colorsLab.GetCurrentFavoritePanel();
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

        private void TestRecentColors(IColorsLabController colorsLab)
        {
            // After all the calls in the earlier tests, recent colors panel should be populated
            IList<Color> currentRecentPanel = colorsLab.GetCurrentRecentPanel();
            AssertEqual(RecentColorsAfterFT, currentRecentPanel);
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



        private static void AssertEqual(IList<Color> expectedList, IList<Color> actualList)
        {
            for (int i = 0; i < expectedList.Count; i++)
            {
                AssertEqual(expectedList[i], actualList[i]);
            }
        }

        private static void AssertEqual(Color expectedColor, Color actualColor)
        {
            // dont assert Alpha
            Assert.IsTrue(IsAlmostSame(expectedColor.R, actualColor.R), "diff color R, expected {0}, actual {1}", expectedColor.R, actualColor.R);
            Assert.IsTrue(IsAlmostSame(expectedColor.G, actualColor.G),"diff color G, expected {0}, actual {1}", expectedColor.G, actualColor.G);
            Assert.IsTrue(IsAlmostSame(expectedColor.B, actualColor.B), "diff color B, expected {0}, actual {1}", expectedColor.B, actualColor.B);
        }

        private static bool IsAlmostSame(byte a, byte b)
        {
            return Math.Abs(a - b) < 3;
        }
        # endregion
    }
}
