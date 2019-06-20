using System;
using System.Windows;
using PowerPointLabs.Utils;
using PowerPointLabs.WPF;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for MessageBox.xaml
    /// </summary>
    public partial class WPFMessageBox
    {
        public enum ButtonPos
        {
            Left,
            Middle,
            Right
        }

        private DialogResult result = Utils.DialogResult.None;

        public WPFMessageBox()
        {
            InitializeComponent();
        }

        public new DialogResult ShowDialog()
        {
            base.ShowDialog();
            return result;
        }

        public void SetButton(ButtonPos pos, DialogResult result)
        {
            MessageButton button = GetButton(pos);
            button.Set(result);
        }

        private MessageButton GetButton(ButtonPos pos)
        {
            switch (pos)
            {
                case ButtonPos.Left:
                    return LeftButton;
                case ButtonPos.Middle:
                    return MiddleButton;
                case ButtonPos.Right:
                    return RightButton;
                default:
                    throw new Exception("Unknown button!");
            }
        }

        private void OnClick_MessageButton(object sender, RoutedEventArgs e)
        {
            MessageButton button = sender as MessageButton;
            if (button == null)
            {
                throw new Exception("Sender button is not a MessageButton!");
            }
            result = button.Result;
            Close();
        }
    }
}
