using System;
using System.Drawing;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

using TestInterface;

using Point = System.Drawing.Point;

namespace Test.FunctionalTest
{
    [TestClass]
    public class ColorsLabTest : BaseFunctionalTest
    {
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

            TestApplyingColors(colorsLab);
            TestRecommendedColors(colorsLab);
            TestFavoriteColors(colorsLab);
            TestColorInfoDialog(colorsLab);
        }

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
        }

        private void TestRecommendedColors(IColorsLabController colorsLab)
        {
            Slide actualSlide = PpOperations.SelectSlide(3);
            PpOperations.SelectShape("selectMe");

            // mono panel7's color will become main color now
            Click(colorsLab.GetMonoPanel7());
            ApplyColor(colorsLab.GetFillColorButton(), colorsLab.GetDropletPanel());

            Slide expSlide = PpOperations.SelectSlide(5);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        private void TestApplyingColors(IColorsLabController colorsLab)
        {
            Slide actualSlide = PpOperations.SelectSlide(3);

            Shape fontColorShape = PpOperations.SelectShape("fontColor")[1];
            Shape lineColorShape = PpOperations.SelectShape("lineColor")[1];
            Shape fillColorShape = PpOperations.SelectShape("fillColor")[1];
            PpOperations.SelectShape("selectMe");

            Panel dropletPanel = colorsLab.GetDropletPanel();

            // get main color from fontColorShape
            // then apply main color to font color of target shape
            ApplyColor(dropletPanel, fontColorShape);
            ApplyColor(colorsLab.GetFontColorButton(), dropletPanel);

            // directly apply font color by fontColorShape's fill color
            ApplyColor(colorsLab.GetLineColorButton(), lineColorShape);

            // get main color from fillColorShape
            // then apply main color to fill color of target shape
            ApplyColor(dropletPanel, fillColorShape);
            ApplyColor(colorsLab.GetFillColorButton(), dropletPanel);

            Slide expSlide = PpOperations.SelectSlide(4);
            SlideUtil.IsSameLooking(expSlide, actualSlide);
        }

        # region Helper methods
        // mouse drag & drop from Control to Shape to apply color
        private void ApplyColor(Control from, Shape to)
        {
            Point startPt = from.PointToScreen(new Point(from.Width/2, from.Height/2));
            Point endPt = new Point(
                PpOperations.PointsToScreenPixelsX(to.Left + to.Width/2),
                PpOperations.PointsToScreenPixelsY(to.Top + to.Height/2));
            DragAndDrop(startPt, endPt);
        }

        // mouse drag & drop from control to another control to apply color
        private void ApplyColor(Control from, Control to)
        {
            Point startPt = from.PointToScreen(new Point(from.Width / 2, from.Height / 2));
            Point endPt = to.PointToScreen(new Point(to.Width / 2, to.Height / 2));
            DragAndDrop(startPt, endPt);
        }

        private void DragAndDrop(Point startPt, Point endPt)
        {
            MouseUtil.SendMouseDown(startPt.X, startPt.Y);
            MouseUtil.SendMouseUp(endPt.X, endPt.Y);
        }

        private void Click(Control target)
        {
            Point pt = target.PointToScreen(new Point(target.Width / 2, target.Height / 2));
            MouseUtil.SendMouseLeftClick(pt.X, pt.Y);
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
