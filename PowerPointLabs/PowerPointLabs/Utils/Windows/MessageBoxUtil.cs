using Forms = System.Windows.Forms;
using MessageBox = PowerPointLabs.WPF.MessageBoxVM;

namespace PowerPointLabs.Utils.Windows
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
            return ShowWinform(text, caption, buttons, icon);
        }

        private static DialogResult ShowWPF(string text, string caption,
            MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            MessageBox messageBox = MessageBox.CreateInstance();
            messageBox.Title = caption;
            messageBox.Message = text;
            messageBox.LeftButton = LeftButtonResult(buttons);
            messageBox.MiddleButton = MiddleButtonResult(buttons);
            messageBox.RightButton = RightButtonResult(buttons);
            messageBox.Icon = icon;
            return messageBox.ShowDialog();
        }

        private static DialogResult ShowWinform(string text, string caption,
            MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return (DialogResult)(int)Forms.MessageBox.Show(
                text, caption, (Forms.MessageBoxButtons)(int)buttons,
                (Forms.MessageBoxIcon)(int)icon);
        }

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
