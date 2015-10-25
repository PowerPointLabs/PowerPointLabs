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
                gotoSlideDialog.Init("Go to the Selected Slide");
                gotoSlideDialog.FocusOkButton();
                this.ShowMetroDialogAsync(gotoSlideDialog, MetroDialogOptions);
            }
            catch
            {
                // dialog could be fired multiple times
            }
        }
    }
}
