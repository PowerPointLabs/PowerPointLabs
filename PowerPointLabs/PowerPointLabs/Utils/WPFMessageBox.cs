using System.Windows.Forms;

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

    public class WPFMessageBox
    {
        public static DialogResult Show(string text)
        {
            return (DialogResult)(int)MessageBox.Show(text);
        }

        public static DialogResult Show(string text, string caption)
        {
            return (DialogResult)(int)MessageBox.Show(text, caption);
        }

        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons)
        {
            return (DialogResult)(int)MessageBox.Show(text, caption, (System.Windows.Forms.MessageBoxButtons)(int)buttons);
        }

        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return (DialogResult)(int)MessageBox.Show(text, caption,
                (System.Windows.Forms.MessageBoxButtons)(int)buttons,
                (System.Windows.Forms.MessageBoxIcon)(int)icon);
        }
    }
}
