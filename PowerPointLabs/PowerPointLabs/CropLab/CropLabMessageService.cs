using System;
using System.Windows;

using PowerPointLabs.CustomControls;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

namespace PowerPointLabs.CropLab
{
    internal class CropLabMessageService : IMessageService
    {
        public void ShowErrorMessageBox(string content, Exception exception = null)
        {
            if (exception != null)
            {
                ErrorDialogBox.ShowDialog(TextCollection.CommonText.ErrorTitle, content, exception);
            }
            else
            {
                WPFMessageBox.Show(content, TextCollection.CommonText.ErrorTitle);
            }
        }
    }
}
