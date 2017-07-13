using System;
using System.Windows;

using PowerPointLabs.CustomControls;
using PowerPointLabs.Views;

namespace PowerPointLabs.CropLab
{
    internal class CropLabMessageService : IMessageService
    {
        public void ShowErrorMessageBox(string content, Exception exception = null)
        {
            if (exception != null)
            {
                ErrorDialogBox.ShowDialog("Error", content, exception);
            }
            else
            {
                MessageBox.Show(content, "Error");
            }
        }
    }
}
