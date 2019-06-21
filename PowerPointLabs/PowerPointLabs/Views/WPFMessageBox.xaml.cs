using System;
using System.Windows;
using PowerPointLabs.WPF;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for MessageBox.xaml
    /// </summary>
    public partial class WPFMessageBox
    {
        public WPFMessageBox()
        {
            InitializeComponent();
        }

        private void OnClick_MessageButton(object sender, RoutedEventArgs e)
        {
            MessageButton button = sender as MessageButton;
            if (button == null)
            {
                System.Windows.Threading.Dispatcher.CurrentDispatcher.InvokeShutdown();
                throw new Exception("Sender button is not a MessageButton!");
            }
            MessageBoxVM data = DataContext as MessageBoxVM;
            if (data == null)
            {
                System.Windows.Threading.Dispatcher.CurrentDispatcher.InvokeShutdown();
                throw new Exception($"Data context is not set to {nameof(MessageBoxVM)}");
            }
            data.Result = button.Result;
            System.Windows.Threading.Dispatcher.CurrentDispatcher.InvokeShutdown();
            Close();
        }
    }
}
