using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.PictureSlidesLab.Model;

namespace Test.UnitTest.PictureSlidesLab.Model
{
    [TestClass]
    public class ObservableImageItemTest
    {
        [TestMethod]
        [TestCategory("UT")]
        public void ImageItemNotification()
        {
            var item = new ObservableImageItem();
            var isNotified = false;
            item.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "ImageItem")
                {
                    isNotified = true;
                }
            };
            item.ImageItem = new ImageItem();
            Assert.IsTrue(isNotified);
        }
    }
}
