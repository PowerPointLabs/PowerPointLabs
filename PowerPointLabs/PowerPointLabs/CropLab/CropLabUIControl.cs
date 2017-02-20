using System;
using System.Windows;

namespace PowerPointLabs.CropLab
{
    internal class CropLabUIControl
    {
        private static CropLabUIControl _instance;

        public static CropLabUIControl GetSharedInstance()
        {
            if (_instance == null)
            {
                _instance = new CropLabUIControl();
            }
            return _instance;
        }

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
