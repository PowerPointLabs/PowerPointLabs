using System;
using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs;
using PowerPointLabs.SyncLab.Views;
using PowerPointLabs.ColorsLab;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Views;
using PowerPointLabs.ShapesLab;

namespace Test.FunctionalTest
{
    [TestClass]
    public class LimitOpenPanesTest : BaseFunctionalTest
    {
        private enum Pane
        {
            ColorsLabPane,
            SyncPane,
            PositionsPane,
            ELearningLabTaskpane,
            CustomShapePane
        }

        private List<Type>[] windowPanes;
        private int activeWindow;

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
        public void FT_LimitOpenPanesTest()
        {
            // start clean
            StartWithNumWindows(2);

            // opening <= 2 panes should work as per normal
            Open(Pane.ColorsLabPane);
            Open(Pane.SyncPane);
            SwitchToWindow(2);

            // Check if old pane is removed
            SwitchToWindow(1);
            Open(Pane.PositionsPane);
            SwitchToWindow(2);

            // Test panes across windows
            Open(Pane.ELearningLabTaskpane);
            Open(Pane.CustomShapePane);
            Open(Pane.PositionsPane);
            SwitchToWindow(1);
        }

        private void Open(Pane PaneType)
        {
            Type paneType = OpenPaneType(PaneType);
            UpdateActivePanes(paneType);
            CheckOpenPanes();
        }

        private void CheckOpenPanes()
        {
            List<Type> expectedWindowPanes = windowPanes[activeWindow];
            HashSet<Type> actualWindowPanes = PpOperations.GetOpenPaneTypes();
            Assert.IsTrue(actualWindowPanes.SetEquals(expectedWindowPanes),
                $"Expected {expectedWindowPanes.Count}, got {actualWindowPanes.Count}.");
        }

        private void UpdateActivePanes(Type paneType)
        {
            List<Type> currentWindowPanes = windowPanes[activeWindow];
            if (currentWindowPanes.Contains(paneType))
            {
                currentWindowPanes.Remove(paneType);
                return;
            }
            currentWindowPanes.Add(paneType);
            PpOperations.SetTagToAssociatedWindow();
            if (currentWindowPanes.Count > 2)
            {
                currentWindowPanes.RemoveAt(0);
            }
        }

        private static Type OpenPaneType(Pane PaneType)
        {
            Type paneType;
            switch (PaneType)
            {
                case Pane.ColorsLabPane:
                    PplFeatures.ColorsLab.OpenPane();
                    paneType = typeof(ColorsLabPane);
                    break;
                case Pane.PositionsPane:
                    PplFeatures.PositionsLab.OpenPane();
                    paneType = typeof(PositionsPane);
                    break;
                case Pane.SyncPane:
                    PplFeatures.SyncLab.OpenPane();
                    paneType = typeof(SyncPane);
                    break;
                case Pane.ELearningLabTaskpane:
                    PplFeatures.ELearningLab.OpenPane();
                    paneType = typeof(ELearningLabTaskpane);
                    break;
                case Pane.CustomShapePane:
                    PplFeatures.ShapesLab.OpenPane();
                    paneType = typeof(CustomShapePane);
                    break;
                default:
                    throw new System.NotImplementedException("Test support for pane not specified!");
            }

            return paneType;
        }

        private void SwitchToWindow(int window)
        {
            if (window < 0 || window >= windowPanes.Length)
            {
                return;
            }
            PpOperations.MaximizeWindow(window);
            activeWindow = window;
            CheckOpenPanes();
        }

        private void StartWithNumWindows(int numWindows)
        {
            // close all but 1 presentation
            while (PpOperations.GetNumWindows() > 1)
            {
                PpOperations.ClosePresentation();
            }

            // create the new windows required
            numWindows = Math.Max(1, numWindows);
            for (int i = 1; i < numWindows; i++)
            {
                PpOperations.NewWindow();
                Assert.AreEqual(PpOperations.GetNumWindows(), numWindows);
            }

            windowPanes = new List<Type>[numWindows + 1];
            for (int i = 1; i <= numWindows; i++)
            {
                windowPanes[i] = new List<Type>();
            }

            // set active window
            SwitchToWindow(1);
        }
    }
}