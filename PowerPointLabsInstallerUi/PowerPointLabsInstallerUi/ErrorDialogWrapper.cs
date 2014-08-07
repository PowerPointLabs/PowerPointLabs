using System;
using System.Windows.Forms;

namespace PowerPointLabs.Views
{
    class ErrorDialogWrapper
    {
        private const string DialogTypeName = "System.Windows.Forms.PropertyGridInternal.GridErrorDlg";

        public static DialogResult ShowDialog(string title, string message, Exception exception)
        {
            var dialogType = typeof(Form).Assembly.GetType(DialogTypeName);

            if (message == string.Empty)
            {
                message = "Something went wrong.";
            }

            if (!message.EndsWith("."))
            {
                message = message + ".";
            }

            // Create dialog instance.
            var dialog = (Form)Activator.CreateInstance(dialogType, new PropertyGrid());
            var completeMsg = message + TextCollection.UserFeedBack + TextCollection.Email;

            // Populate relevant properties on the dialog instance.
            dialog.Text = title;
            dialogType.GetProperty("Details").SetValue(dialog,
                                                       exception.GetType() + "\r\n" +
                                                       exception.Message + "\r\n" +
                                                       "Stack Trace:\r\n" + 
                                                       exception.StackTrace,
                                                       null);
            dialogType.GetProperty("Message").SetValue(dialog, completeMsg, null);

            return dialog.ShowDialog();
        }
    }
}
