using System;
using System.Windows;

using PowerPointLabs.CustomControls;

namespace PowerPointLabs.CropLab
{
    internal class CropLabMessageService : IMessageService
    {
        public void ShowErrorMessageBox(string content, Exception exception = null)
        {
            if (exception != null)
            {
                Views.ErrorDialogWrapper.ShowDialog("Error", content, exception);
            }
            else
            {
                MessageBox.Show(content, "Error");
            }
        }
    }
}
