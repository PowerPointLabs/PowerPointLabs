using System.Windows.Forms;

namespace PowerPointLabs.Utils
{
    //public enum DialogResult
    //{
    //    None = 0,
    //    OK = 1,
    //    Cancel = 2,
    //    Abort = 3,
    //    Retry = 4,
    //    Ignore = 5,
    //    Yes = 6,
    //    No = 7
    //}

    //public enum MessageBoxButtons
    //{
    //    OK = 0,
    //    OKCancel = 1,
    //    AbortRetryIgnore = 2,
    //    YesNoCancel = 3,
    //    YesNo = 4,
    //    RetryCancel = 5
    //}

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
            return (DialogResult)(int)MessageBox.Show(text, caption, MessageBoxButtons.AbortRetryIgnore);
        }

        public static DialogResult Show(string text, string caption, MessageBoxButtons buttons, MessageBoxIcon icon)
        {
            return (DialogResult)(int)MessageBox.Show(text, caption, buttons, icon);
        }
    }
}
