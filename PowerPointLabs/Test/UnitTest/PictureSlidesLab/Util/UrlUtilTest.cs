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

        private string _googleImgLinkThatNeedDecode =
            "https://www.google.com.sg/imgres?imgurl=" +
            "http://www.virgin.com/sites/default/files/Articles/Entrepreneur%252520Getty/Entrepreneur_breakfast_getty.jpg&" +
            "imgrefurl=http://www.virgin.com/entrepreneur/in-focus-the-rise-of-flexible-working&h=1415&w=2122&" +
            "tbnid=eNWrKlo61EFchM:&docid=OEazz6Oo0s_77M&ei=5n2cVuLIJY_juQTuoJbIAQ&" +
            "tbm=isch&ved=0ahUKEwji0-CA1rLKAhWPcY4KHW6QBRkQMwhQKBYwFg";

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
            ImageItem imgItem = new ImageItem();
            string link = _googleImgLink.Clone() as string;
            UrlUtil.GetMetaInfo(ref link, imgItem);
            Assert.AreEqual("http://tctechcrunch2011.files.wordpress.com/2011/05/tcdisrupt_tc-9.jpg",
                link);
            Assert.AreEqual("http://techcrunch.com/2011/05/21/the-hack-is-on-at-the-hackathon/",
                imgItem.ContextLink);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestGetMetaInfoWithDecoding()
        {
            ImageItem imgItem = new ImageItem();
            string link = _googleImgLinkThatNeedDecode.Clone() as string;
            UrlUtil.GetMetaInfo(ref link, imgItem);
            Assert.AreEqual("http://www.virgin.com/sites/default/files/Articles/Entrepreneur%20Getty/Entrepreneur_breakfast_getty.jpg",
                link);
            Assert.AreEqual("http://www.virgin.com/entrepreneur/in-focus-the-rise-of-flexible-working",
                imgItem.ContextLink);
        }
    }
}
