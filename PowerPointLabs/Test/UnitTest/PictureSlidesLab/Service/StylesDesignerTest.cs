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
        private StyleOptionsFactory _factory;

        protected override string GetTestingSlideName()
        {
            return "PictureSlidesLab\\StylesDesigner.pptx";
        }

        [TestInitialize]
        public void Init()
        {
            _factory = new StyleOptionsFactory();
            _designer = new StylesDesigner(App);
            _sourceImage = new ImageItem
            {
                ImageFile = PathUtil.GetDocTestPath() + "PictureSlidesLab\\koala.jpg",
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
            foreach (StyleOption style in _factory.GetAllStylesPreviewOptions())
            {
                PowerPointLabs.PictureSlidesLab.Service.Preview.PreviewInfo previewInfo = _designer.PreviewApplyStyle(_sourceImage, _contentSlide,
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
            foreach (StyleOption style in _factory.GetAllStylesPreviewOptions())
            {
                _designer.ApplyStyle(_sourceImage, _contentSlide,
                    Pres.PageSetup.SlideWidth, Pres.PageSetup.SlideHeight, style);
                string imgPath = TempPath.GetPath("applystyle-" +
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
