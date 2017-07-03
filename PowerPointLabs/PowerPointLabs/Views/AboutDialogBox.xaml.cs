using System;
using System.Windows;
using System.Windows.Navigation;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for AboutDialogBox.xaml
    /// </summary>
    public partial class AboutDialogBox
    {
        public AboutDialogBox(string versionNumber, string releaseDate, string websiteUrl)
        {
            InitializeComponent();

            versionRunText.Text = versionNumber;
            if (Properties.Settings.Default.ReleaseType == "dev")
            {
                versionRunText.Text += " (Dev-release)";
            }

            releaseDateRunText.Text = releaseDate;

            websiteHyperlink.NavigateUri = new Uri(websiteUrl);
            websiteHyperlinkRunText.Text = websiteUrl;
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
