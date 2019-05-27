using System;
using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab.ObjectFormats;
using PowerPointLabs.SyncLab.Views;

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

            TestFollowingNodes(nodes1);
            TestFollowingNodes(nodes2);
            TestFollowingNodes(nodes3);
        }

        /// <summary>
        /// Get the node names of all format nodes.
        /// </summary>
        /// <returns></returns>
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
                GetEndNode(formats, name).IsChecked = true;
            }
            return formats;
        }

        /// <summary>
        /// Gets the node at the end of the format tree, as specified by the string name
        /// Returns early with the last matching node if the next node cannot be found
        /// </summary>
        private FormatTreeNode GetEndNode(FormatTreeNode[] formats, string name)
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

        /// <summary>
        /// Test the node names displayed by the tooltip against the expected node names.
        /// </summary>
        private void TestFollowingNodes(string[] nodeNames)
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
