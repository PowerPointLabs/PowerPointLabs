using System.Windows;
using MahApps.Metro.Controls.Dialogs;

namespace PowerPointLabs.ImageSearch
{
    partial class ImageSearchPane
    {
        private void GotoSlideButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                _gotoSlideDialog.Init("Go to the Selected Slide");
                _gotoSlideDialog.FocusOkButton();
                this.ShowMetroDialogAsync(_gotoSlideDialog, MetroDialogOptions);
            }
            catch
            {
                // dialog could be fired multiple times
            }
        }
    }
}
