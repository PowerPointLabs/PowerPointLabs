using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;

namespace Test.UnitTest.ELearningLab.Model
{
    [TestClass]
    public class ClickItemTest
    {
        private ClickItem item;

        [TestInitialize]
        public void Init()
        {
            item = new ClickItem();
        }

        [TestMethod]
        [TestCategory("UT")]
        public void ClickNoChangedNotification()
        {
            bool notified = false;
            item.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "ClickNo")
                {
                    notified = true;
                }
            };
            item.ClickNo = 1;
            Assert.IsTrue(notified);
        }

    }
}
