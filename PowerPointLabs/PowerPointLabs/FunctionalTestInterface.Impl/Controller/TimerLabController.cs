using System;
using System.Windows;
using System.Windows.Controls.Primitives;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.TimerLab;
using TestInterface;

namespace PowerPointLabs.FunctionalTestInterface.Impl.Controller
{
    [Serializable]
    class TimerLabController : MarshalByRefObject, ITimerLabController
    {
        private static ITimerLabController _instance = new TimerLabController();

        public static ITimerLabController Instance { get { return _instance; } }

        private TimerPane _pane;

        private TimerLabController() {}

        public void OpenPane()
        {
            UIThreadExecutor.Execute(() =>
            {
                FunctionalTestExtensions.GetRibbonUi().OnAction(
                    new RibbonControl("TimerLabButton"));
                _pane = FunctionalTestExtensions.GetTaskPane(
                    typeof(TimerPane)).Control as TimerPane;
            });
        }

        public void ClickCreateButton()
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    _pane.TimerPaneWPF.createButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                });
            }
        }

        public void SetHeightTextBoxValue(int value)
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    _pane.TimerPaneWPF.HeightTextBox.Text = "" + value;
                    // set focus to the other text box to trigger the change
                    _pane.TimerPaneWPF.WidthTextBox.Focus();
                });
            }
        }

        public int GetHeightTextBoxValue()
        {
            int value = -1;

            if (_pane != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    value = Int32.Parse(_pane.TimerPaneWPF.HeightTextBox.Text);
                });
            }
            return value;
        }

        public void SetWidthTextBoxValue(int value)
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    _pane.TimerPaneWPF.WidthTextBox.Text = "" + value;
                    // set focus to the other text box to trigger the change
                    _pane.TimerPaneWPF.HeightTextBox.Focus();
                });
            }
        }

        public int GetWidthTextBoxValue()
        {
            int value = -1;

            if (_pane != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    value = Int32.Parse(_pane.TimerPaneWPF.WidthTextBox.Text);
                });
            }
            return value;
        }
    }
}
