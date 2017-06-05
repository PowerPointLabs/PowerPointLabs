using System;
using System.Reflection;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.SyncLab;
using PowerPointLabs.SyncLab.ObjectFormats;

namespace Test.UnitTest.SyncLab
{
    [TestClass]
    public class SyncLabFormatTest
    {
        [TestMethod]
        [TestCategory("UT")]
        public void SyncLabValidFormats()
        {
            FormatTreeNode[] formatCategories = SyncFormatConstants.FormatCategories;
            AssertValidFormat(formatCategories);
        }

        private void AssertValidFormat(FormatTreeNode[] nodes)
        {
            foreach (FormatTreeNode node in nodes)
            {
                AssertValidFormat(node);
            }
        }

        private void AssertValidFormat(FormatTreeNode node)
        {
            if (node.IsFormatNode)
            {
                AssertValidFormatType(node.Format.FormatType);
            }
            else
            {
                AssertValidFormat(node.ChildrenNodes);
            }
        }

        private void AssertValidFormatType(Type type)
        {
            MethodInfo method = type.GetMethod("CanCopy", BindingFlags.Public | BindingFlags.Static);
            Assert.IsNotNull(method);
            method = type.GetMethod("SyncFormat", BindingFlags.Public | BindingFlags.Static);
            Assert.IsNotNull(method);
            method = type.GetMethod("DisplayImage", BindingFlags.Public | BindingFlags.Static);
            Assert.IsNotNull(method);
        }
    }
}
