using System;
using System.ComponentModel;
using System.Threading;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

namespace PowerPointLabs.WPF
{
    public class MessageBoxVM : BaseNotificationClass, IWPFWindow, INotifyPropertyChanged
    {
        public string Title { get; set; }
        public string Message { get; set; }
        public DialogResult LeftButton { get; set; }
        public DialogResult MiddleButton { get; set; }
        public DialogResult RightButton { get; set; }
        public string IconSource { get; set; }
        public DialogResult Result { get; set; }
        public bool Visible
        {
            get
            {
                return _visible;
            }
            set
            {
               _visible = value;
                if (_visible)
                {
                    closed.Reset();
                }
                else
                {
                    closed.Set();
                }
            }
        }
        private bool _visible = false;
        private ManualResetEventSlim closed = new ManualResetEventSlim(false);

        public static MessageBoxVM CreateInstance()
        {
            WPFMessageBox messageBox = new WPFMessageBox();
            MessageBoxVM data = new MessageBoxVM();
            messageBox.DataContext = data;
            return data;
        }

        // Obsolete
        public void SetIcon(MessageBoxIcon icon)
        {
            IconSource = GetIcon(icon);
            OnPropertyChanged(nameof(IconSource));
        }

        public DialogResult ShowDialog()
        {
            Result = DialogResult.None;
            Visible = true;
            closed.Wait();
            return Result;
        }

        private string GetIcon(MessageBoxIcon icon)
        {
            switch (icon)
            {
                case MessageBoxIcon.Asterisk:
                    return "..\\Resources\\About.png";
                case MessageBoxIcon.Error:
                    return "..\\Resources\\Help.png";
                case MessageBoxIcon.Exclamation:
                    return "..\\Resources\\Help.png";
                case MessageBoxIcon.Question:
                    return "..\\Resources\\Help.png";
                case MessageBoxIcon.None:
                default:
                    return "";
            }
        }
    }
}
