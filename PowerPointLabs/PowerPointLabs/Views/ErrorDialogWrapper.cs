using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    class ErrorDialogWrapper
    {
        private const string userFeedBack =
            @"You are highly appreciated if you can email the error details to us for further improving.";

        private const string email = @"pptlabs@comp.nus.edu.sg";

        public static DialogResult showDialog(string title, string message, Exception exception)
        {
            var dialogTypeName = "System.Windows.Forms.PropertyGridInternal.GridErrorDlg";
            var dialogType = typeof(Form).Assembly.GetType(dialogTypeName);

            // Create dialog instance.
            var dialog = (Form)Activator.CreateInstance(dialogType, new PropertyGrid());

            var completeMsg = message + "\n" + userFeedBack + "\n" + email;

            // Populate relevant properties on the dialog instance.
            dialog.Text = title;
            dialogType.GetProperty("Details").SetValue(dialog,
                                                       exception.Message + "\nStack Trace:\n" + exception.StackTrace,
                                                       null);
            dialogType.GetProperty("Message").SetValue(dialog, completeMsg, null);

            return dialog.ShowDialog();
        }
    }
}
