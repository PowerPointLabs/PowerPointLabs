using System;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.ModelFactory;
using PowerPointLabs.PictureSlidesLab.Service;
using PowerPointLabs.PictureSlidesLab.Util;
using Test.Util;

namespace Test.UnitTest.PictureSlidesLab.Service
{
    [TestClass]
    public class StylesDesignerTest : BaseUnitTest
    {
        private StylesDesigner _designer;
        private ImageItem _sourceImage;
        private Slide _contentSlide;

        protected override string GetTestingSlideName()
        {
            return "PictureSlidesLab\\StylesDesigner.pptx";
        }

        [TestInitialize]
        public void Init()
        {
            _designer = new StylesDesigner(App);
            _sourceImage = new ImageItem
            {
                ImageFile = PathUtil.GetDocTestPath() + "koala.jpg",
                Tooltip = "some tooltip"
            };
            _contentSlide = PpOperations.SelectSlide(1);
        }

        [TestCleanup]
        public void CleanUp()
        {
            _designer.CleanUp();
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestPreviewStyle()
        {
            TempPath.InitTempFolder();
            foreach (var style in StyleOptionsFactory.GetAllStylesPreviewOptions())
            {
                var previewInfo = _designer.PreviewApplyStyle(_sourceImage, _contentSlide,
                    Pres.PageSetup.SlideWidth, Pres.PageSetup.SlideHeight, style);
                SlideUtil.IsSameLooking(
                    new FileInfo(PathUtil.GetDocTestPath() +
                        "PictureSlidesLab\\" +
                        style.StyleName + ".jpg"),
                    new FileInfo(previewInfo.PreviewApplyStyleImagePath));
            }
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestApplyStyle()
        {
            TempPath.InitTempFolder();
            foreach (var style in StyleOptionsFactory.GetAllStylesPreviewOptions())
            {
                _designer.ApplyStyle(_sourceImage, _contentSlide,
                    Pres.PageSetup.SlideWidth, Pres.PageSetup.SlideHeight, style);
                var imgPath = TempPath.GetPath("applystyle-" +
                                               Guid.NewGuid().ToString().Substring(0, 7) +
                                               "-" + DateTime.Now.GetHashCode());
                _contentSlide.Export(imgPath, "JPG");
                SlideUtil.IsSameLooking(
                    new FileInfo(PathUtil.GetDocTestPath() +
                        "PictureSlidesLab\\" +
                        style.StyleName + ".jpg"),
                    new FileInfo(imgPath));
            }
        }
    }
}
