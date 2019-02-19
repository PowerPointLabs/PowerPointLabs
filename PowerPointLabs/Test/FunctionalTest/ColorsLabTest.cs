using System;
using System.Drawing;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

using TestInterface;

using Point = System.Windows.Point;
using Button = System.Windows.Controls.Button;


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

        private const string TargetShape = "selectMe";
        private const string FontColorShape = "fontColor";
        private const string LineColorShape = "lineColor";
        private const string FillColorShape = "fillColor";

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

            //TestRecommendedColors(colorsLab);
            //TestFavoriteColors(colorsLab);
            //TestColorInfoDialog(colorsLab);
        }

        private void TestApplyFontColor(IColorsLabController colorsLab)
        {
            Slide actualSlide = PpOperations.SelectSlide(OriginalSlideNo);

            Shape targetShape = PpOperations.SelectShape(FontColorShape)[1];
            Point startPt = colorsLab.GetApplyTextButtonLocation();
            Point endPt = GetShapeCenterPoint(targetShape);

            PpOperations.SelectShape(TargetShape);
            DragAndDrop(startPt, endPt);

            Slide expSlide = PpOperations.SelectSlide(FontColorChangeSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void TestApplyLineColor(IColorsLabController colorsLab)
        {
            Slide actualSlide = PpOperations.SelectSlide(OriginalSlideNo);

            Shape targetShape = PpOperations.SelectShape(LineColorShape)[1];
            Point startPt = colorsLab.GetApplyLineButtonLocation();
            Point endPt = GetShapeCenterPoint(targetShape);

            PpOperations.SelectShape(TargetShape);
            DragAndDrop(startPt, endPt);

            Slide expSlide = PpOperations.SelectSlide(OutlineColorChangeSlideNo);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void TestApplyFillColor(IColorsLabController colorsLab)
        {
            Slide actualSlide = PpOperations.SelectSlide(OriginalSlideNo);

            Shape targetShape = PpOperations.SelectShape(FillColorShape)[1];
            Point startPt = colorsLab.GetApplyFillButtonLocation();
            Point endPt = GetShapeCenterPoint(targetShape);

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

        /*
        private void TestColorInfoDialog(IColorsLabController colorsLab)
        {
            IColorsLabMoreInfoDialog infoDialog = null;
            try
            {
                infoDialog = colorsLab.ShowMoreColorInfo(colorsLab.GetMonoPanel1().BackColor);
                // rgb text is like "RGB: 163, 192, 242"
                string[] rgbColor = infoDialog.GetRgbText().Substring(5).Split(',');
                int r = Int32.Parse(rgbColor[0].Trim());
                int g = Int32.Parse(rgbColor[1].Trim());
                int b = Int32.Parse(rgbColor[2].Trim());
                // rgb values can have errors within threshold 2
                Assert.IsTrue(Math.Abs(r - 163) <= 2);
                Assert.IsTrue(Math.Abs(g - 192) <= 2);
                Assert.IsTrue(Math.Abs(b - 242) <= 2);
            }
            finally
            {
                if (infoDialog != null) infoDialog.TearDown();
            }
        }

        private void TestFavoriteColors(IColorsLabController colorsLab)
        {
            Panel favPanel1 = colorsLab.GetFavColorPanel1();
            Color originalFavColor = favPanel1.BackColor;

            try
            {
                // empty fav colors
                colorsLab.GetEmptyFavColorsButton().PerformClick();
                Color colorAftReset = favPanel1.BackColor;
                AssertEqual(Color.White, colorAftReset);

                // set mono panel1's color as fav color
                Panel monoPanel1 = colorsLab.GetMonoPanel1();
                ApplyColor(monoPanel1, favPanel1);
                AssertEqual(monoPanel1.BackColor, favPanel1.BackColor);
            }
            finally
            {
                // reset fav colors from last time saved
                colorsLab.GetResetFavColorsButton().PerformClick();
                AssertEqual(originalFavColor, favPanel1.BackColor);
            }
        
            */

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


        private void ApplyColor(System.Windows.Controls.Control from, System.Windows.Controls.Control to)
        {
            Point startPt = from.PointToScreen(new Point(from.Width / 2, from.Height / 2));
            Point endPt = to.PointToScreen(new Point(to.Width / 2, to.Height / 2));
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

        /*
        private void Click(Control target)
        {
            Point pt = target.PointToScreen(new Point(target.Width / 2, target.Height / 2));
            MouseUtil.SendMouseLeftClick(pt.X, pt.Y);
        }*/

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
