using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointLabs
{
    class ErrorDialogWrapper
    {
        public static void showDialog(string title, string message, string details)
        {
            var dialogTypeName = "System.Windows.Forms.PropertyGridInternal.GridErrorDlg";
            var dialogType = typeof(Form).Assembly.GetType(dialogTypeName);

            // Create dialog instance.
            var dialog = (Form)Activator.CreateInstance(dialogType, new PropertyGrid());

            // Populate relevant properties on the dialog instance.
            dialog.Text = title;
            dialogType.GetProperty("Details").SetValue(dialog, details, null);
            dialogType.GetProperty("Message").SetValue(dialog, message, null);

            var result = dialog.ShowDialog();
        }
    }
}
