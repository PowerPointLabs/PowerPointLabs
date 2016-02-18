using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;

namespace PowerPointLabs.ResizeLab
{
    public partial class ResizeLabPaneWPF
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
