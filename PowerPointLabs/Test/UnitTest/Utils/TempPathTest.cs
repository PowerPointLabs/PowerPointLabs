using System.IO;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.Utils;

namespace Test.UnitTest.Utils
{
    [TestClass]
    public class TempPathTest
    {
        [TestMethod]
        [TestCategory("UT")]
        public void TestTempTestFolderName()
        {
            string tempFolder = TempPath.GetTempTestFolder();
            string folderName = Path.GetFileName(Path.GetDirectoryName(tempFolder));
            Assert.AreEqual("PowerPointLabsTest", folderName);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestCreateTempTestFolder()
        {
            TempPath.CreateTempTestFolder();
            Assert.IsTrue(TempPath.IsExistingTempTestFolder());
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDeleteTempTestFolder()
        {
            TempPath.CreateTempTestFolder();
            TempPath.DeleteTempTestFolder();
            Assert.IsFalse(TempPath.IsExistingTempTestFolder());
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDeleteTempTestFolderWithReadOnlyFile()
        {
            string readOnlyFile = Path.Combine(TempPath.GetTempTestFolder(), "ReadOnlyFile.txt");

            TempPath.CreateTempTestFolder();

            using (File.Create(readOnlyFile))
            {
                File.SetAttributes(readOnlyFile, FileAttributes.ReadOnly);
            }

            TempPath.DeleteTempTestFolder();
            Assert.IsFalse(TempPath.IsExistingTempTestFolder());
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestDeleteTempTestFolderWithNestedFolder()
        {
            string subFolder = Path.Combine(TempPath.GetTempTestFolder(), "SubPowerPointLabsTest\\");

            TempPath.CreateTempTestFolder();

            if (!Directory.Exists(subFolder))
            {
                Directory.CreateDirectory(subFolder);
            }

            TempPath.DeleteTempTestFolder();
            Assert.IsFalse(TempPath.IsExistingTempTestFolder());
        }

        [TestCleanup]
        public void TearDown()
        {
            TempPath.CreateTempTestFolder();
        }
    }
}
