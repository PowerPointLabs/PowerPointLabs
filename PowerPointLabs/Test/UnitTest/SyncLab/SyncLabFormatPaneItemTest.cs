using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.SyncLab.ObjectFormats;
using PowerPointLabs.SyncLab.Views;
using System;
using System.Collections.Generic;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabFormatPaneItemTest : BaseSyncLabTest
    {
        private const string TestName = "TestingN@meeeeeeeeeeeeeeeeee";

        private string[] _testNodesNames;

        [TestInitialize]
        public void TestInitialize()
        {
            _testNodesNames = GetAllFormatNames();
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncPaneTooltipName()
        {
            FormatTreeNode[] nodes = SyncFormatConstants.FormatCategories;
            SyncFormatPaneItem item = new SyncFormatPaneItemStub(nodes);
            item.Text = TestName;

            Assert.AreEqual(TestName, item.toolTipName.Text);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestSyncPaneTooltipBody()
        {
            string[] nodes1 = { _testNodesNames[0], _testNodesNames[13], _testNodesNames[16] };
            string[] nodes2 = { _testNodesNames[4], _testNodesNames[6], _testNodesNames[8] };
            string[] nodes3 = { _testNodesNames[7], _testNodesNames[11], _testNodesNames[18] };

            testFollowingNodes(nodes1);
            testFollowingNodes(nodes2);
            testFollowingNodes(nodes3);
        }

        private string[] GetAllFormatNames()
        {
            Queue<FormatTreeNode> queue = new Queue<FormatTreeNode>(SyncFormatConstants.FormatCategories);
            List<string> allNames = new List<string>();
            while (queue.Count > 0)
            {
                FormatTreeNode node = queue.Dequeue();
                string fullName = node.Name;
                FormatTreeNode parentNode = node.ParentNode;
                while (parentNode != null)
                {
                    fullName = parentNode.Name + SyncFormatConstants.FormatNameSeparator + fullName;
                    parentNode = parentNode.ParentNode;
                }
                allNames.Add(fullName);

                if (node.ChildrenNodes == null)
                {
                    continue;
                }
                foreach (FormatTreeNode child in node.ChildrenNodes)
                {
                    queue.Enqueue(child);
                }
            }

            return allNames.ToArray();
        }

        private FormatTreeNode[] SelectFormats(FormatTreeNode[] formats, string[] formatNames)
        {
            foreach (string name in formatNames)
            {
                getEndNode(formats, name).IsChecked = true;
            }
            return formats;
        }

        private FormatTreeNode getEndNode(FormatTreeNode[] formats, string name)
        {
            string[] path = name.Split(SyncFormatConstants.FormatNameSeparator.ToCharArray());
            FormatTreeNode currentRoot = new FormatTreeNode("root", formats);
            foreach (string nodeName in path)
            {
                int nextIndex = LastIndexOf(currentRoot.ChildrenNodes, nodeName);
                if (nextIndex < 0)
                {
                    break;
                }
                currentRoot = currentRoot.ChildrenNodes[nextIndex];
            }
            return currentRoot;
        }

        private int LastIndexOf(FormatTreeNode[] nodes, string name)
        {
            int index = nodes.Length - 1;
            while (index >= 0)
            {
                if (nodes[index].Name.Equals(name))
                {
                    return index;
                }
                index--;
            }
            return index;
        }

        private void testFollowingNodes(string[] nodeNames)
        {
            FormatTreeNode[] nodes = SyncFormatConstants.FormatCategories;
            nodes = SelectFormats(SyncFormatConstants.FormatCategories, nodeNames);
            SyncFormatPaneItem item = new SyncFormatPaneItemStub(nodes);

            string[] expected = nodeNames;
            string[] actual = item.toolTipBody.Text.Split("\n".ToCharArray());
            Array.Sort(expected);
            Array.Sort(actual);

            Assert.AreEqual(String.Join("\n", expected), String.Join("\n", actual));
        }
    }
}
