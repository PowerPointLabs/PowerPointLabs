using ImageProcessor.Imaging.Filters;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Service;
using PowerPointLabs.PictureSlidesLab.Service.Effect;
using PowerPointLabs.PictureSlidesLab.Util;

using Test.Util;

namespace Test.UnitTest.PictureSlidesLab.Service
{
    [TestClass]
    public class EffectsDesignerTest : BaseUnitTest
    {
        public readonly string Img = PathUtil.GetDocTestPath() + "PictureSlidesLab\\koala.jpg";
        public readonly string Link = "http://www.google.com/";

        private Slide _contentSlide;
        private Slide _processingSlide;
        private EffectsDesigner _designer;
        private ImageItem _imgItem;

        protected override string GetTestingSlideName()
        {
            return "PictureSlidesLab\\EffectsDesigner.pptx";
        }

        [TestInitialize]
        public void Init()
        {
            _contentSlide = PpOperations.SelectSlide(1);
            _processingSlide = PpOperations.SelectSlide(2);
            _imgItem = new ImageItem
            {
                ImageFile = Img,
                Tooltip = "some tooltips"
            };
            _designer = new EffectsDesigner(_contentSlide,
                Pres.PageSetup.SlideWidth, Pres.PageSetup.SlideHeight,
                _imgItem);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestInsertImageReference()
        {
            // constructor for producing preview image
            EffectsDesigner ed = new EffectsDesigner(_processingSlide);
            ed.PreparePreviewing(
                _contentSlide, 
                Pres.PageSetup.SlideWidth, Pres.PageSetup.SlideHeight, 
                new ImageItem
                {
                    ImageFile = "some images",
                    Tooltip = "some tooltips"
                });

            ed.ApplyImageReferenceToSlideNote(Link);
            Assert.IsTrue( 
                PpOperations.GetNotesPageText(_processingSlide)
                .Contains(Link));

            ed.ApplyImageReferenceInsertion(Link, "Calibri", "#000000", 14, "", Alignment.Left);
            Microsoft.Office.Interop.PowerPoint.ShapeRange refShape = PpOperations.SelectShapesByPrefix(
                EffectsDesigner.ShapeNamePrefix + "_" + EffectName.ImageReference);
            Assert.IsTrue(
                refShape.TextFrame2.TextRange.Text.Contains(Link));
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestInsertBackground()
        {
            Microsoft.Office.Interop.PowerPoint.Shape bgShape = _designer.ApplyBackgroundEffect();
            Assert.IsTrue(bgShape.Name.StartsWith(
                EffectsDesigner.ShapeNamePrefix + "_" + EffectName.BackGround));
            Assert.AreEqual(MsoShapeType.msoPicture, bgShape.Type);
            Assert.AreEqual(0f, bgShape.Left);
            Assert.AreEqual(0f, bgShape.Top);
            Assert.AreEqual(540f, bgShape.Height);
            Assert.AreEqual(960f, bgShape.Width);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestTextEffect()
        {
            _designer.ApplyTextEffect("Tahoma", "#123456", 3, 0);
            PpOperations.SelectSlide(1);
            Microsoft.Office.Interop.PowerPoint.ShapeRange shape = PpOperations.SelectShapesByPrefix("Title");
            float originalTextSize = float.Parse(shape.Tags[Tag.OriginalFontSize]);

            Assert.AreEqual("Tahoma", shape.TextEffect.FontName);
            Assert.AreEqual(originalTextSize + 3, shape.TextEffect.FontSize);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestTextPositionAndAlignment()
        {
            _designer.ApplyTextPositionAndAlignment(Position.Left, Alignment.Auto);
            TextBoxInfo tbInfo = new TextBoxes(
                _contentSlide.Shapes.Range(), 
                Pres.PageSetup.SlideWidth, 
                Pres.PageSetup.SlideHeight)
                .GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(25f, tbInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(177.52f, tbInfo.Top));
            Assert.IsTrue(SlideUtil.IsRoughlySame(640.76f, tbInfo.Width));
            Assert.IsTrue(SlideUtil.IsRoughlySame(184.96f, tbInfo.Height));

            _designer.ApplyTextWrapping();
            tbInfo = new TextBoxes(
                _contentSlide.Shapes.Range(),
                Pres.PageSetup.SlideWidth,
                Pres.PageSetup.SlideHeight)
                .GetTextBoxesInfo();
            Assert.IsTrue(SlideUtil.IsRoughlySame(25f, tbInfo.Left));
            Assert.IsTrue(SlideUtil.IsRoughlySame(119.2f, tbInfo.Top));
            Assert.IsTrue(SlideUtil.IsRoughlySame(418.505035f, tbInfo.Width));
            Assert.IsTrue(SlideUtil.IsRoughlySame(243.279984f, tbInfo.Height));
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestOverlayEffect()
        {
            Microsoft.Office.Interop.PowerPoint.Shape shape = _designer.ApplyOverlayEffect("#000000", 35);
            Assert.IsTrue(shape.Name.StartsWith(
                EffectsDesigner.ShapeNamePrefix + "_" + EffectName.Overlay));
            Assert.AreEqual(MsoShapeType.msoAutoShape, shape.Type);
            Assert.AreEqual(0f, shape.Left);
            Assert.AreEqual(0f, shape.Top);
            Assert.AreEqual(540f, shape.Height);
            Assert.AreEqual(960f, shape.Width);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestBlurEffect()
        {
            TempPath.InitTempFolder();
            Microsoft.Office.Interop.PowerPoint.Shape shape = _designer.ApplyBlurEffect();

            Assert.IsTrue(shape.Name.StartsWith(
                EffectsDesigner.ShapeNamePrefix + "_" + EffectName.Blur));
            Assert.AreEqual(MsoShapeType.msoPicture, shape.Type);
            Assert.IsNotNull(_imgItem.BlurImageFile);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestTextBoxEffect()
        {
            _designer.ApplyTextboxEffect("#000000", 35, 0);
            PpOperations.SelectSlide(1);
            Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange = PpOperations.SelectShapesByPrefix(
                EffectsDesigner.ShapeNamePrefix + "_" + EffectName.TextBox);

            Assert.AreEqual(1, shapeRange.Count);
            Assert.AreEqual(MsoShapeType.msoAutoShape, shapeRange[1].Type);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestCircleRingsEffect()
        {
            Microsoft.Office.Interop.PowerPoint.Shape shape = _designer.ApplyCircleRingsEffect("#000000", 35);
            Assert.IsTrue(shape.Name.StartsWith(
                EffectsDesigner.ShapeNamePrefix + "_" + EffectName.Overlay));
            Assert.AreEqual(MsoShapeType.msoGroup, shape.Type);
            Assert.AreEqual(2, shape.GroupItems.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestRectBannerEffect()
        {
            Microsoft.Office.Interop.PowerPoint.Shape shape = _designer.ApplyRectBannerEffect(BannerDirection.Auto, Position.Left,
                null, "#000000", 35);
            Assert.IsTrue(shape.Name.StartsWith(
                EffectsDesigner.ShapeNamePrefix + "_" + EffectName.Banner));
            Assert.AreEqual(MsoShapeType.msoAutoShape, shape.Type);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestRectOutlineEffect()
        {
            Microsoft.Office.Interop.PowerPoint.Shape shape = _designer.ApplyRectOutlineEffect(null, "#000000", 35);
            Assert.IsTrue(shape.Name.StartsWith(
                EffectsDesigner.ShapeNamePrefix + "_" + EffectName.Banner));
            Assert.AreEqual(MsoShapeType.msoAutoShape, shape.Type);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestFrameEffect()
        {
            Microsoft.Office.Interop.PowerPoint.Shape shape = _designer.ApplyAlbumFrameEffect("#000000", 35);
            Assert.IsTrue(shape.Name.StartsWith(
                EffectsDesigner.ShapeNamePrefix + "_" + EffectName.Overlay));
            Assert.AreEqual(MsoShapeType.msoAutoShape, shape.Type);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestTriangleEffect()
        {
            Microsoft.Office.Interop.PowerPoint.Shape shape = _designer.ApplyTriangleEffect("#FFFFFF", "#000000", 35);
            Assert.IsTrue(shape.Name.StartsWith(
                EffectsDesigner.ShapeNamePrefix + "_" + EffectName.Overlay));
            Assert.AreEqual(MsoShapeType.msoGroup, shape.Type);
            Assert.AreEqual(2, shape.GroupItems.Count);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSpecialEffectsEffect()
        {
            TempPath.InitTempFolder();
            Microsoft.Office.Interop.PowerPoint.Shape shape = _designer.ApplySpecialEffectEffect(MatrixFilters.GreyScale, true);
            Assert.IsTrue(shape.Name.StartsWith(
                EffectsDesigner.ShapeNamePrefix + "_" + EffectName.SpecialEffect));
            Assert.AreEqual(MsoShapeType.msoPicture, shape.Type);
            Assert.IsNotNull(_imgItem.SpecialEffectImageFile);
        }
    }
}
