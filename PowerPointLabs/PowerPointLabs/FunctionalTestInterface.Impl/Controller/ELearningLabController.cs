using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Views;
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.TextCollection;

using TestInterface;

namespace PowerPointLabs.FunctionalTestInterface.Impl.Controller
{
    [Serializable]
    class ELearningLabController : MarshalByRefObject, IELearningLabController
    {
        private static IELearningLabController _instance = new ELearningLabController();

        public static IELearningLabController Instance { get { return _instance; } }

        private ELearningLabTaskpane _pane;

        private ELearningLabController() { }

        public void OpenPane()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                FunctionalTestExtensions.GetRibbonUi().OnAction(
                    new RibbonControl(ELearningLabText.ELearningTaskPaneTag));
                _pane = FunctionalTestExtensions.GetTaskPane(
                    typeof(ELearningLabTaskpane)).Control as ELearningLabTaskpane;
            }));
        }

        public void AddSelfExplanationItem()
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    _pane.ELearningLabMainPanel.createButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                    ExplanationItem item1 = _pane.ELearningLabMainPanel.listView.Items
                        .OfType<ClickItem>().Last() as ExplanationItem;
                    item1.CaptionText = "Test self explanation item 1";
                    item1.IsCaption = true;
                    _pane.ELearningLabMainPanel.createButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                    ExplanationItem item2 = _pane.ELearningLabMainPanel.listView.Items
                    .OfType<ClickItem>().Last() as ExplanationItem;
                    item2.CaptionText = "Test self explanation item 2";
                    item2.CalloutText = "This is a shorter callout for self explanation item 2";
                    item2.IsShortVersionIndicated = true;
                    item2.IsCaption = true;
                    item2.IsCallout = true;
                    _pane.ELearningLabMainPanel.createButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                    ExplanationItem item3 = _pane.ELearningLabMainPanel.listView.Items
                    .OfType<ClickItem>().Last() as ExplanationItem;
                    item3.CaptionText = "Test self explanation item 3";
                    item3.IsCaption = true;
                    item3.IsCallout = true;
                }));
            }
        }

        public void Sync()
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    _pane.ELearningLabMainPanel.syncButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));                    
                }));
            }
        }

        public void Reorder()
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    ListViewItem item = (ListViewItem)_pane.ELearningLabMainPanel.listView
                    .ItemContainerGenerator.ContainerFromIndex(0);
                    Button downButton = VisualTreeUtility.FindByName("downButton", item) as Button;
                    if (downButton != null)
                    {
                        downButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                    }
                }));
            }
        }

        public void Delete()
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    ListViewItem item = (ListViewItem)_pane.ELearningLabMainPanel.listView
                    .ItemContainerGenerator.ContainerFromIndex(0);
                    Button deleteButton = VisualTreeUtility.FindByName("deleteButton", item) as Button;
                    if (deleteButton != null)
                    {
                        deleteButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                    }
                }));
            }
        }
    }
}
