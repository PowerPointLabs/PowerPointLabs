using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PictureSlidesLab.Service.Effect;

using Test.Util;

namespace Test.UnitTest.PictureSlidesLab.Service.Effect
{
    [TestClass]
    public class TextBoxesTest : BaseUnitTest
    {
        protected override string GetTestingSlideName()
        {
            return "PictureSlidesLab\\TextBoxes.pptx";
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTextBoxInfoForEmptyTextBoxes()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = PpOperations.SelectShapesByPrefix("TextBox");
            TextBoxes textBoxes = new TextBoxes(shapes, 
                Pres.PageSetup.SlideWidth, Pres.PageSetup.SlideHeight);
            TextBoxInfo textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.AreEqual(null, textBoxInfo);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetTextBoxInfo()
        {
            PpOperations.SelectSlide(2);
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = PpOperations.SelectShapesByPrefix("TextBox");
            TextBoxes textBoxes = new TextBoxes(shapes,
                Pres.PageSetup.SlideWidth, Pres.PageSetup.SlideHeight);
            TextBoxInfo textBoxInfo = textBoxes.GetTextBoxesInfo();

            Assert.IsTrue(SlideUtil.IsRoughlySame(57.5200043f, textBoxInfo.Height));
            Assert.IsTrue(SlideUtil.IsRoughlySame(68.2f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(52.17752f, textBoxInfo.Top));
            Assert.IsTrue(SlideUtil.IsRoughlySame(710.945068f, textBoxInfo.Width));

            TextBoxes.AddMargin(textBoxInfo, 25);
            Assert.IsTrue(SlideUtil.IsRoughlySame(107.520004f, textBoxInfo.Height));
            Assert.IsTrue(SlideUtil.IsRoughlySame(43.1999969f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(27.17752f, textBoxInfo.Top));
            Assert.IsTrue(SlideUtil.IsRoughlySame(760.945068f, textBoxInfo.Width));
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStartBoxing()
        {
            PpOperations.SelectSlide(2);
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = PpOperations.SelectShapesByPrefix("TextBox");
            TextBoxes textBoxes = new TextBoxes(shapes,
                Pres.PageSetup.SlideWidth, Pres.PageSetup.SlideHeight);

            textBoxes
                .SetAlignment(Alignment.Centre)
                .SetPosition(Position.Left)
                .StartBoxing();
            TextBoxInfo textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(57.5200043f, textBoxInfo.Height));
            Assert.IsTrue(SlideUtil.IsRoughlySame(25f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(241.240036f, textBoxInfo.Top));
            Assert.IsTrue(SlideUtil.IsRoughlySame(710.945068f, textBoxInfo.Width));

            textBoxes
                .SetAlignment(Alignment.Centre)
                .SetPosition(Position.Centre)
                .StartBoxing();
            textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(124.527481f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(241.240036f, textBoxInfo.Top));

            textBoxes
                .SetAlignment(Alignment.Centre)
                .SetPosition(Position.BottomLeft)
                .StartBoxing();
            textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(25f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(457.480042f, textBoxInfo.Top));

            textBoxes
                .SetAlignment(Alignment.Centre)
                .SetPosition(Position.Bottom)
                .StartBoxing();
            textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(124.527481f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(457.480042f, textBoxInfo.Top));

            textBoxes
                .SetAlignment(Alignment.Centre)
                .SetPosition(Position.Right)
                .StartBoxing();
            textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(224.054962f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(241.23996f, textBoxInfo.Top));
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStartBoxingWithTextWrapping()
        {
            PpOperations.SelectSlide(2);
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapes = PpOperations.SelectShapesByPrefix("TextBox");
            TextBoxes textBoxes = new TextBoxes(shapes,
                Pres.PageSetup.SlideWidth, Pres.PageSetup.SlideHeight);

            textBoxes.StartTextWrapping();

            textBoxes
                .SetAlignment(Alignment.Centre)
                .SetPosition(Position.Left)
                .StartBoxing();

            TextBoxInfo textBoxInfo = textBoxes.GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(105.040009f, textBoxInfo.Height));
            Assert.IsTrue(SlideUtil.IsRoughlySame(25.00004f, textBoxInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(217.480042f, textBoxInfo.Top));
            // aft text wrapping, width is smaller (originally should be 710)
            Assert.IsTrue(SlideUtil.IsRoughlySame(365.47f, textBoxInfo.Width));
        }
    }
}
