using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Navigation;

namespace PowerPointLabs.ImageSearch
{
    /// <summary>
    /// Interaction logic for SearchOptionsPane.xaml
    /// </summary>
    public partial class SearchOptionsPane
    {
        public SearchOptionsPane()
        {
            InitializeComponent();
            DataContextChanged += (sender, args) =>
            {
                ChangeVisibilityForSearchEngineOptions();
            };
        }

        private void Hyperlink_OnRequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

        private void SearchEngineComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ChangeVisibilityForSearchEngineOptions();
        }

        private void ChangeVisibilityForSearchEngineOptions()
        {
            Dispatcher.BeginInvoke(new Action(() => { 
                if (SearchEngineComboBox.SelectedIndex == 0 /*Bing*/)
                {
                    ChangeVisibilityForBingOptions(Visibility.Visible);
                    ChangeVisibilityForGoogleOptions(Visibility.Hidden);
                }
                else
                {
                    ChangeVisibilityForBingOptions(Visibility.Hidden);
                    ChangeVisibilityForGoogleOptions(Visibility.Visible);
                }
            }));
        }

        private void ChangeVisibilityForGoogleOptions(Visibility vis)
        {
            GoogleSearchEngineIdLabel.Visibility = vis;
            GoogleSearchEngineIdTextBox.Visibility = vis;
            
            GoogleApiKeyLabel.Visibility = vis;
            GoogleApiKeyTextBox.Visibility = vis;

            GoogleColorTypeLabel.Visibility = vis;
            GoogleColorTypeComboBox.Visibility = vis;

            GoogleDominantColorLabel.Visibility = vis;
            GoogleDominantColorComboBox.Visibility = vis;

            GoogleImageTypeLabel.Visibility = vis;
            GoogleImageTypeComboBox.Visibility = vis;

            GoogleImageSizeLabel.Visibility = vis;
            GoogleImageSizeComboBox.Visibility = vis;

            GoogleFileTypeLabel.Visibility = vis;
            GoogleFileTypeComboBox.Visibility = vis;
        }

        private void ChangeVisibilityForBingOptions(Visibility vis)
        {
            BingSearchEngineIdLabel.Visibility = vis;
            BingSearchEngineIdTextBox.Visibility = vis;

            BingImageSizeLabel.Visibility = vis;
            BingImageSizeComboBox.Visibility = vis;

            BingImageColorLabel.Visibility = vis;
            BingImageColorComboBox.Visibility = vis;

            BingImageStyleLabel.Visibility = vis;
            BingImageStyleComboBox.Visibility = vis;

            BingImageFaceLabel.Visibility = vis;
            BingImageFaceComboBox.Visibility = vis;
        }
    }
}
