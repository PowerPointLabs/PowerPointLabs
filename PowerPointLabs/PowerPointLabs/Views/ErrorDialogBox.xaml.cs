using System;
using System.Windows;
using System.Windows.Navigation;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for ErrorDialogBox.xaml
    /// </summary>
    public partial class ErrorDialogBox
    {
        public static void ShowDialog(string title, string message, Exception exception)
        {
            ErrorDialogBox dialog = new ErrorDialogBox(title, message, exception);
            dialog.ShowThematicDialog();
        }

        private ErrorDialogBox(string title, string message, Exception exception)
        {
            InitializeComponent();

            titleText.Text = title;

            if (message == string.Empty)
            {
                message = "Something went wrong.";
            }

            if (!message.EndsWith("."))
            {
                message = message + ".";
            }

            errorMessageText.Text = message;

            emailHyperlink.NavigateUri = new Uri("mailto:" + CommonText.ReportIssueEmail);
            emailHyperlinkRunText.Text = CommonText.ReportIssueEmail;

            fullMessageText.Text = exception.GetType() + "\r\n" +
                                    exception.Message + "\r\n" +
                                    "Stack Trace:\r\n" +
                                    exception.StackTrace;
        }

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            System.Diagnostics.Process.Start(e.Uri.ToString());
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
