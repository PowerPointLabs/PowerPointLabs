using System.Collections.Generic;
using System.Drawing;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PictureSlidesLab.Util;

using Test.Util;

namespace Test.UnitTest.PictureSlidesLab.Util
{
    [TestClass]
    public class ImageUtilTest
    {
        [TestMethod]
        [TestCategory("UT")]
        public void TestGetThumbnailFromFullSizeImg()
        {
            StoragePath.InitPersistentFolder();
            StoragePath.CleanPersistentFolder(new List<string>());
            string thumbnail = 
                ImageUtil.GetThumbnailFromFullSizeImg(
                    PathUtil.GetDocTestPath() + "PictureSlidesLab\\koala.jpg");
            Bitmap thumbnailImage = new Bitmap(thumbnail);
            Bitmap fullsizeImage = new Bitmap(
                PathUtil.GetDocTestPath() + "PictureSlidesLab\\koala.jpg");
            Assert.IsTrue(thumbnailImage.Width < fullsizeImage.Width 
                && thumbnailImage.Height < fullsizeImage.Height);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetWidthAndHeight()
        {
            string result = ImageUtil.GetWidthAndHeight(
                PathUtil.GetDocTestPath() + "PictureSlidesLab\\koala.jpg");
            Assert.AreEqual("500 x 375", result);
        }
    }
}
