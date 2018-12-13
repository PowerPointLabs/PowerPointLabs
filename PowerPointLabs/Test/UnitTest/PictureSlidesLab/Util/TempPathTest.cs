using System;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PictureSlidesLab.Util;

namespace Test.UnitTest.PictureSlidesLab.Util
{
    [TestClass]
    public class TempPathTest
    {
        [TestMethod]
        [TestCategory("UT")]
        public void TestTempPathInit()
        {
            try
            {
                TempPath.GetPath("some-name");
                Assert.Fail();
            }
            catch(Exception)
            {
            }
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestTempPathInit2()
        {
            try
            {
                TempPath.InitTempFolder();
                TempPath.GetPath("some-name");
            }
            catch (Exception)
            {
                Assert.Fail();
            }
        }
    }
}
