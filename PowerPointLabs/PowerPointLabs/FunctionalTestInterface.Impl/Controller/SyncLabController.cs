using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.SyncLab.Views;
using PowerPointLabs.TextCollection;

using TestInterface;

namespace PowerPointLabs.FunctionalTestInterface.Impl.Controller
{
    [Serializable]
    class SyncLabController : MarshalByRefObject, ISyncLabController
    {
        private static ISyncLabController _instance = new SyncLabController();

        public static ISyncLabController Instance { get { return _instance; } }

        private SyncPane _pane;

        private SyncLabController() { }

        public void OpenPane()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                FunctionalTestExtensions.GetRibbonUi().OnAction(
                    new RibbonControl(SyncLabText.PaneTag));
                _pane = FunctionalTestExtensions.GetTaskPane(
                    typeof(SyncPane)).Control as SyncPane;
            }));
        }

        public void Copy()
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    _pane.SyncPaneWPF1.copyButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                });
            }
        }

        public void Sync(int index)
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    ((SyncFormatPaneItem)_pane.SyncPaneWPF1.formatListBox.Items[index])
                            .pasteButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                });
            }
        }

        public void DialogSelectItem(int categoryIndex, int itemIndex)
        {
            if (_pane != null && _pane.SyncPaneWPF1.Dialog != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    ((_pane.SyncPaneWPF1.Dialog.treeView.Items[categoryIndex] as TreeViewItem)
                        .Items[itemIndex] as SyncFormatDialogItem).IsChecked = true;
                });
            }
        }

        public void DialogClickOk()
        {
            if (_pane != null && _pane.SyncPaneWPF1.Dialog != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    _pane.SyncPaneWPF1.Dialog.okButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                });
            }
            
        }
    
        public bool GetCopyButtonEnabledStatus()
        {
            bool result = false;
            UIThreadExecutor.Execute(() =>
            {
                result = _pane.SyncPaneWPF1.GetCopyButtonEnabledStatus();
            });
            return result;
        }
    }
}
