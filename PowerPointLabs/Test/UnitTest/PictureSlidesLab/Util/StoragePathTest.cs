using System;
using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.PictureSlidesLab.Util;

namespace Test.UnitTest.PictureSlidesLab.Util
{
    [TestClass]
    public class StoragePathTest
    {
        [TestMethod]
        [TestCategory("UT")]
        public void TestStoragePathInit()
        {
            try
            {
                StoragePath.GetPath("some-name");
                Assert.Fail();
            }
            catch (Exception)
            {
            }
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestStoragePathInit2()
        {
            try
            {
                StoragePath.InitPersistentFolder();
                StoragePath.CleanPersistentFolder(new List<string>());
                StoragePath.GetPath("some-name");
            }
            catch (Exception)
            {
                Assert.Fail();
            }
        }
    }
}
