using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Views;
using PowerPointLabs.ELearningLab.Service;
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

        public string DefaultVoiceLabel => string.Format(ELearningLabText.AudioDefaultLabelFormat, AudioSettingService.selectedVoice.VoiceName);

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

        public void CreateTemplateExplanations(params ExplanationItemTemplate[] items)
        {
            // add to IELearningLabController
            if (_pane != null)
            {
                _pane.ELearningLabMainPanel.Dispatcher.Invoke((Action)(() =>
                {
                    _pane.ELearningLabMainPanel.Items.Clear();
                    foreach (ExplanationItemTemplate item in items)
                    {
                        ExplanationItem explanationItem = CreateExplanationItem();
                        explanationItem.CopyFormat(item);
                        if (!explanationItem.HasSameFormat(item))
                        {
                            throw new Exception("Format failed to copy correctly");
                        }
                    }
                }));
            }
        }

        public ExplanationItemTemplate[] GetExplanations()
        {
            return _pane.ELearningLabMainPanel.Dispatcher.Invoke(() =>
            {
                List<ExplanationItemTemplate> result = new List<ExplanationItemTemplate>();
                foreach (ExplanationItem item in _pane.ELearningLabMainPanel.Items.OfType<ExplanationItem>())
                {
                    ExplanationItemTemplate template = new ExplanationItemTemplate();
                    template.CopyFormat(item);
                    result.Add(template);
                }
                return result.ToArray();
            });
        }

        public void AddAbove(int index)
        {
            _pane.ELearningLabMainPanel.Dispatcher.Invoke(() =>
            {
                ListViewItem visual = _pane.ELearningLabMainPanel.listView.ItemContainerGenerator.ContainerFromIndex(index) as ListViewItem;
                if (visual == null)
                {
                    throw new Exception("Null visual when retrieving list item");
                }
                MenuItem item = visual.ContextMenu.Items.OfType<MenuItem>().First(
                    o => (string)o.Header == "Add Explanation Above");
                if (item == null)
                {
                    throw new Exception("Null menu item");
                }
                item.CommandParameter = _pane.ELearningLabMainPanel.listView.Items.GetItemAt(index);
                item.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
            });
        }

        public void AddBelow(int index)
        {
            _pane.ELearningLabMainPanel.Dispatcher.Invoke(() =>
            {
                ListViewItem visual = _pane.ELearningLabMainPanel.listView.ItemContainerGenerator.ContainerFromIndex(index) as ListViewItem;
                if (visual == null)
                {
                    throw new Exception("Null visual when retrieving list item");
                }
                MenuItem item = visual.ContextMenu.Items.OfType<MenuItem>().First(
                    o => (string)o.Header == "Add Explanation Below");
                if (item == null)
                {
                    throw new Exception("Null menu item");
                }
                item.CommandParameter = _pane.ELearningLabMainPanel.listView.Items.GetItemAt(index);
                item.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
            });
        }

        public void AddAtBottom()
        {
            _pane.ELearningLabMainPanel.Dispatcher.Invoke(() =>
            {
                CreateExplanationItem();
            });
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
