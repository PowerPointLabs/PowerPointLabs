using System;
using System.Windows;
using System.Windows.Controls.Primitives;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.SyncLab.View;
using TestInterface;

namespace PowerPointLabs.FunctionalTestInterface.Impl.Controller
{
    [Serializable]
    class SyncLabController : MarshalByRefObject, ISyncLabController
    {
        private static ISyncLabController _instance = new SyncLabController();

        public static ISyncLabController Instance { get { return _instance; } }

        private SyncPane _pane;
        private SyncFormatDialog _dialog;

        private SyncLabController() { }

        public void OpenPane()
        {
            UIThreadExecutor.Execute(() =>
            {
                FunctionalTestExtensions.GetRibbonUi().OnAction(
                    new RibbonControl("SyncLabButton"));
                _pane = FunctionalTestExtensions.GetTaskPane(
                    typeof(SyncPane)).Control as SyncPane;
            });
        }

        public void OpenCopyDialog()
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    _pane.SyncPaneWPF1.copyButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                    _dialog = FunctionalTestExtensions.GetTaskPane(
                        typeof(SyncFormatDialog)).Control as SyncFormatDialog;
                });
            }
        }

        public void Copy()
        {
            OpenCopyDialog();
            if (_dialog != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    _dialog.okButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                    _dialog = null;
                });
            }
        }

        public void Sync()
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    //_pane.positionsPaneWPF.duplicateRotationButton.IsChecked = !_pane.positionsPaneWPF.duplicateRotationButton.IsChecked;
                });
            }
        }


    }
}
