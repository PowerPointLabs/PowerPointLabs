using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.SyncLab.ObjectFormats;
using PowerPointLabs.SyncLab.Views;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabFormatPaneItemTest : BaseSyncLabTest
    {
        private const int OriginalShapesSlideNo = 4;
        private const string CopyFromShape = "CopyFrom";
        private const string TestName = "TestingNameeeeeeeeeeeeeeeeee";

        private Shape _formatShape;

        [TestInitialize]
        public void TestInitialize()
        {
            _formatShape = GetShape(OriginalShapesSlideNo, CopyFromShape);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncPaneTooltip()
        {
            FormatTreeNode[] nodes = SyncFormatConstants.FormatCategories;
            SyncFormatPaneItem item = new SyncFormatPaneItemStub(nodes);
            item.Text = TestName;

            Assert.AreEqual(TestName, item.toolTipName.Text);
            Assert.AreEqual("", item.toolTipBody.Text);
        }

    }
}
