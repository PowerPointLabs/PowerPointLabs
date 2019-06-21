using System.Windows.Controls;

using PowerPointLabs.Utils.Windows;

namespace PowerPointLabs.WPF
{
    public class MessageButton : Button
    {
        public DialogResult Result { get; private set; }
        public void Set(DialogResult result)
        {
            this.Result = result;
            if (result == DialogResult.None)
            {
                Visibility = System.Windows.Visibility.Hidden;
            }
            switch (result)
            {
                case DialogResult.Abort:
                    Content = nameof(DialogResult.Abort);
                    break;
                case DialogResult.Cancel:
                    Content = nameof(DialogResult.Cancel);
                    break;
                case DialogResult.Ignore:
                    Content = nameof(DialogResult.Ignore);
                    break;
                case DialogResult.No:
                    Content = nameof(DialogResult.No);
                    break;
                case DialogResult.None:
                    Content = nameof(DialogResult.None);
                    break;
                case DialogResult.OK:
                    Content = nameof(DialogResult.OK);
                    break;
                case DialogResult.Retry:
                    Content = nameof(DialogResult.Retry);
                    break;
                case DialogResult.Yes:
                    Content = nameof(DialogResult.Yes);
                    break;
                default:
                    break;
            }
        }
    }
}
