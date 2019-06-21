using System;
using System.ComponentModel;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using PowerPointLabs.Utils.Windows;
using PowerPointLabs.Views;

namespace PowerPointLabs.WPF
{
    public class MessageBoxVM : BaseNotificationClass, IWPFWindow, INotifyPropertyChanged
    {
        public string Title
        {
            get
            {
                return _title;
            }
            set
            {
                _title = value;
                OnPropertyChanged(nameof(Title));
            }
        }
        public string Message
        {
            get
            {
                return _message;
            }
            set
            {
                _message = value;
                OnPropertyChanged(nameof(Message));
            }
        }
        public DialogResult LeftButton
        {
            get
            {
                return _leftButton;
            }
            set
            {
                _leftButton = value;
                OnPropertyChanged(nameof(LeftButton));
            }
        }
        public DialogResult MiddleButton
        {
            get
            {
                return _middleButton;
            }
            set
            {
                _middleButton = value;
                OnPropertyChanged(nameof(MiddleButton));
            }
        }
        public DialogResult RightButton
        {
            get
            {
                return _rightButton;
            }
            set
            {
                _rightButton = value;
                OnPropertyChanged(nameof(RightButton));
            }
        }
        public MessageBoxIcon Icon
        {
            get
            {
                return _icon;
            }
            set
            {
                _icon = value;
                OnPropertyChanged(nameof(Icon));
            }
        }
        public DialogResult Result
        {
            get
            {
                return _result;
            }
            set
            {
                _result = value;
                OnPropertyChanged(nameof(Result));
            }
        }

        public Visibility Visible
        {
            get
            {
                return _visible;
            }
            set
            {
               _visible = value;
                if (_visible == Visibility.Visible)
                {
                    closed.Reset();
                }
                else
                {
                    closed.Set();
                }
                OnPropertyChanged(nameof(Visible));
            }
        }

        private string _title;
        private string _message;
        private DialogResult _leftButton;
        private DialogResult _middleButton;
        private DialogResult _rightButton;
        private MessageBoxIcon _icon;
        private DialogResult _result = DialogResult.None;
        private Visibility _visible = Visibility.Hidden;
        private ManualResetEventSlim closed = new ManualResetEventSlim(false);

        public static MessageBoxVM CreateInstance()
        {
            WPFMessageBox messageBox = new WPFMessageBox();
            MessageBoxVM data = new MessageBoxVM();
            messageBox.DataContext = data;
            return data;
        }

        private MessageBoxVM()
        {

        }

        public DialogResult ShowDialog()
        {
            Result = DialogResult.None;
            Visible = Visibility.Visible;
            Dispatcher.Run();
            //closed.Wait(); No need for this anymore
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
