using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Util;

namespace Test.UnitTest.PictureSlidesLab.Util
{
    [TestClass]
    public class UrlUtilTest
    {
        private string _googleImgLink =
            "https://www.google.com.sg/imgres?imgurl=" +
            "http://tctechcrunch2011.files.wordpress.com/2011/05/tcdisrupt_tc-9.jpg" +
            "&imgrefurl=http://techcrunch.com/2011/05/21/the-hack-is-on-at-the-hackathon/" +
            "&h=1284&w=1920&tbnid=FhLTqsOxEDgXMM:&docid=A03pP_VlVkpIXM&ei=fwSVVrn4A4a40gTzspDwBQ" +
            "&tbm=isch&ved=0ahUKEwj5s--etaTKAhUGnJQKHXMZBF4QMwg3KAcwBw";

        [TestMethod]
        [TestCategory("UT")]
        public void TestIsUrlValid()
        {
            Assert.IsTrue(UrlUtil.IsUrlValid("http://google.com"));
            Assert.IsFalse(UrlUtil.IsUrlValid("google.com"));
            Assert.IsFalse(UrlUtil.IsUrlValid(""));
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestIsValidGoogleImageLink()
        {
            Assert.IsTrue(UrlUtil.IsUrlValid(_googleImgLink));
            Assert.IsFalse(UrlUtil.IsUrlValid("google.com"));
            Assert.IsFalse(UrlUtil.IsUrlValid(""));
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetMetaInfo()
        {
            var imgItem = new ImageItem();
            var link = _googleImgLink.Clone() as string;
            UrlUtil.GetMetaInfo(ref link, imgItem);
            Assert.AreEqual("http://tctechcrunch2011.files.wordpress.com/2011/05/tcdisrupt_tc-9.jpg",
                link);
            Assert.AreEqual("http://techcrunch.com/2011/05/21/the-hack-is-on-at-the-hackathon/",
                imgItem.ContextLink);
        }
    }
}
