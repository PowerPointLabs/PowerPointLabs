using System;
using System.Windows.Media;
using MessageBox = PowerPointLabs.Views.WPFMessageBox;

namespace PowerPointLabs.Utils
{
    public enum DialogResult
    {
        None = 0,
        OK = 1,
        Cancel = 2,
        Abort = 3,
        Retry = 4,
        Ignore = 5,
        Yes = 6,
        No = 7
    }

    public enum MessageBoxButtons
    {
        OK = 0,
        OKCancel = 1,
        AbortRetryIgnore = 2,
        YesNoCancel = 3,
        YesNo = 4,
        RetryCancel = 5
    }

    public enum MessageBoxIcon
    {
        None = 0,
        Hand = 16,
        Stop = 16,
        Error = 16,
        Question = 32,
        Exclamation = 48,
        Warning = 48,
        Asterisk = 64,
        Information = 64
    }

    public class MessageBoxUtil
    {
        public static DialogResult Show(string text, string caption = "",
            MessageBoxButtons buttons = MessageBoxButtons.OK, MessageBoxIcon icon = MessageBoxIcon.None)
        {
            MessageBox messageBox = new MessageBox();
            messageBox.Title = caption;
            messageBox.Message.Text = text;
            messageBox.SetButton(MessageBox.ButtonPos.Left, LeftButtonResult(buttons));
            messageBox.SetButton(MessageBox.ButtonPos.Middle, MiddleButtonResult(buttons));
            messageBox.SetButton(MessageBox.ButtonPos.Right, RightButtonResult(buttons));
            // Set the image source!
            //messageBox.Icon = GetImageSource(icon);
            return messageBox.CustomShowDialog();
        }

        //private static ImageSource GetImageSource(MessageBoxIcon icon)
        //{
        //    switch (icon)
        //    {
        //        case MessageBoxIcon.Asterisk:
        //            break;
        //        case MessageBoxIcon.None:
        //        default:
        //            break;
        //    }
        //}

        private static DialogResult LeftButtonResult(MessageBoxButtons buttons)
        {
            switch (buttons)
            {
                case MessageBoxButtons.YesNo:
                case MessageBoxButtons.YesNoCancel:
                    return DialogResult.Yes;
                case MessageBoxButtons.AbortRetryIgnore:
                    return DialogResult.Abort;
                default:
                    return DialogResult.None;
            }
        }
        private static DialogResult MiddleButtonResult(MessageBoxButtons buttons)
        {
            switch (buttons)
            {
                case MessageBoxButtons.OKCancel:
                    return DialogResult.OK;
                case MessageBoxButtons.RetryCancel:
                case MessageBoxButtons.AbortRetryIgnore:
                    return DialogResult.Retry;
                case MessageBoxButtons.YesNo:
                    return DialogResult.Yes;
                case MessageBoxButtons.YesNoCancel:
                    return DialogResult.No;
                default:
                    return DialogResult.None;
            }
        }

        private static DialogResult RightButtonResult(MessageBoxButtons buttons)
        {
            switch (buttons)
            {
                case MessageBoxButtons.AbortRetryIgnore:
                    return DialogResult.Ignore;
                case MessageBoxButtons.OK:
                    return DialogResult.OK;
                case MessageBoxButtons.YesNo:
                    return DialogResult.No;
                case MessageBoxButtons.OKCancel:
                case MessageBoxButtons.RetryCancel:
                case MessageBoxButtons.YesNoCancel:
                    return DialogResult.Cancel;
                default:
                    return DialogResult.None;
            }
        }
    }
}
