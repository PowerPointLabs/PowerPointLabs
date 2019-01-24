using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Util;

using Test.Util;

namespace Test.UnitTest.PictureSlidesLab.Model
{
    [TestClass]
    public class ImageItemTest
    {
        private ImageItem item;

        [TestInitialize]
        public void Init()
        {
            item = new ImageItem();
        }

        [TestMethod]
        [TestCategory("UT")]
        public void ImageFileNotification()
        {
            bool notified = false;
            item.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "ImageFile")
                {
                    notified = true;
                }
            };
            item.ImageFile = "something";
            Assert.IsTrue(notified);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void ToolTipNotification()
        {
            bool notified = false;
            item.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "Tooltip")
                {
                    notified = true;
                }
            };
            item.Tooltip = "something";
            Assert.IsTrue(notified);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestUpdateDownloadedImage()
        {
            StoragePath.InitPersistentFolder();
            StoragePath.CleanPersistentFolder(new List<string>());
            string imagePath = PathUtil.GetDocTestPath() + "PictureSlidesLab\\koala.jpg";
            item.UpdateDownloadedImage(imagePath);
            Assert.AreEqual(imagePath, item.FullSizeImageFile);
            Assert.IsFalse(string.IsNullOrEmpty(item.ImageFile));
            Assert.AreEqual("500 x 375", item.Tooltip);
        }
    }
}
