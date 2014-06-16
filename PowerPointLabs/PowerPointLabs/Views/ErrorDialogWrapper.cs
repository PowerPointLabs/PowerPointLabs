using System;
using System.Text;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    class ErrorDialogWrapper
    {
        private const string UserFeedBack =
            @" Help us fix the problem by emailing ";

        private const string Email = @"pptlabs@comp.nus.edu.sg";

        private const string DialogTypeName = "System.Windows.Forms.PropertyGridInternal.GridErrorDlg";

        public static DialogResult ShowDialog(string title, string message, Exception exception)
        {
            var dialogType = typeof(Form).Assembly.GetType(DialogTypeName);

            // Create dialog instance.
            var dialog = (Form)Activator.CreateInstance(dialogType, new PropertyGrid());
            var completeMsg = message + UserFeedBack + Email;

            // Populate relevant properties on the dialog instance.
            dialog.Text = title;
            dialogType.GetProperty("Details").SetValue(dialog,
                                                       exception.Message + "\nStack Trace:\n" + exception.StackTrace,
                                                       null);
            dialogType.GetProperty("Message").SetValue(dialog, completeMsg.ToString(), null);

            return dialog.ShowDialog();
        }
    }
}
