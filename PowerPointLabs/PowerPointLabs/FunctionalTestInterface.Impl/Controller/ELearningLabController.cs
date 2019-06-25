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

        public void CreateTemplateExplanations(params IExplanationItem[] items)
        {
            // add to IELearningLabController
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    _pane.ELearningLabMainPanel.Items.Clear();
                }));
                foreach (IExplanationItem item in items)
                {
                    ExplanationItem explanationItem = CreateExplanationItem();
                    item.CopyFormat(item);
                    if (!explanationItem.HasSameFormat(item))
                    {
                        throw new Exception("Format failed to copy correctly");
                    }
                }
            }
        }

        public void AddSelfExplanationItem()
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    ExplanationItem item1 = CreateExplanationItem();
                    item1.CaptionText = "Test self explanation item 1";
                    item1.IsCaption = true;
                    ExplanationItem item2 = CreateExplanationItem();
                    item2.CaptionText = "Test self explanation item 2";
                    item2.CalloutText = "This is a shorter callout for self explanation item 2";
                    item2.IsCaption = true;
                    item2.IsCallout = true;
                    ExplanationItem item3 = CreateExplanationItem();
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

        private ExplanationItem CreateExplanationItem()
        {
            _pane.ELearningLabMainPanel.createButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
            return _pane.ELearningLabMainPanel.listView.Items
            .OfType<ClickItem>().Last() as ExplanationItem;
        }
    }
}
