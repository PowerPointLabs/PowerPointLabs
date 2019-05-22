using System;
using System.Collections.Generic;

using Microsoft.Office.Tools;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs;
using Test.Util;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.FunctionalTestInterface.Impl;
using System.Threading;
using PowerPointLabs.SyncLab.Views;
using PowerPointLabs.ColorsLab;
using System.Windows.Forms;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Views;
using PowerPointLabs.ShapesLab;

namespace Test.FunctionalTest
{
    [TestClass]
    public class LimitOpenPanesTest : BaseFunctionalTest
    {

        protected override bool IsUseNewPpInstance()
        {
            return true;
        }

        protected override string GetTestingSlideName()
        {
            // any slide can be used
            return "ShortcutsLab\\QuickProperties.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AATest()
        {
            // start clean
            int windows = PpOperations.GetNumWindows();
            PpOperations.NewWindow();
            Assert.AreEqual(PpOperations.GetNumWindows(), windows + 1);
            PpOperations.MaximizeWindow(1);
            // should work as per normal
            HashSet<Type> window1Panes = new HashSet<Type>();
            HashSet<Type> window2Panes = new HashSet<Type>();
            PplFeatures.ColorsLab.OpenPane();
            window1Panes.Add(typeof(ColorsLabPane));
            Assert.IsTrue(PpOperations.GetOpenPaneTypes().SetEquals(window1Panes));
            PplFeatures.SyncLab.OpenPane();
            window1Panes.Add(typeof(SyncPane));
            Assert.IsTrue(PpOperations.GetOpenPaneTypes().SetEquals(window1Panes));

            PpOperations.MaximizeWindow(2);
            // this is required to activate the window for the correct openpanetypes to show up
            Assert.IsTrue(PpOperations.GetOpenPaneTypes().SetEquals(window2Panes));

            // Check if old pane is removed
            // TODO: Get API to close the panes
            PpOperations.MaximizeWindow(1);
            PplFeatures.PositionsLab.OpenPane();
            window1Panes.Add(typeof(PositionsPane));
            window1Panes.Remove(typeof(ColorsLabPane));
            Assert.IsTrue(PpOperations.GetOpenPaneTypes().SetEquals(window1Panes));

            PpOperations.MaximizeWindow(2);
            Assert.IsTrue(PpOperations.GetOpenPaneTypes().SetEquals(window2Panes));

            // TODO: Touch the other window
            // Somehow the number of panes associated with the second window is zero
            /*
            PplFeatures.ELearningLab.OpenPane();
            window2Panes.Add(typeof(ELearningLabTaskpane));
            Assert.IsTrue(PpOperations.GetOpenPaneTypes().SetEquals(window2Panes));
            PplFeatures.ShapesLab.OpenPane();
            window2Panes.Add(typeof(CustomShapePane));
            Assert.IsTrue(PpOperations.GetOpenPaneTypes().SetEquals(window2Panes));
            PplFeatures.PositionsLab.OpenPane();
            window2Panes.Add(typeof(PositionsPane));
            window2Panes.Remove(typeof(ELearningLabTaskpane));
            Assert.IsTrue(PpOperations.GetOpenPaneTypes().SetEquals(window2Panes));

            PpOperations.MaximizeWindow(1);
            Assert.IsTrue(PpOperations.GetOpenPaneTypes().SetEquals(window1Panes));
            */
        }
    }
}