using System.Windows.Media;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PictureSlidesLab.Model;

namespace Test.UnitTest.PictureSlidesLab.Model
{
    [TestClass]
    public class ObservableFontTest
    {
        [TestMethod]
        [TestCategory("UT")]
        public void FontNotification()
        {
            ObservableFont font = new ObservableFont();
            bool isNotified = false;
            font.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "Font")
                {
                    isNotified = true;
                }
            };
            font.Font = new FontFamily("");
            Assert.IsTrue(isNotified);
        }
    }
}
