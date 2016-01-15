using System.Reflection;
using System.Windows;
using MahApps.Metro.Controls.Dialogs;
using PowerPointLabs.Models;

namespace PowerPointLabs.PictureSlidesLab.View
{
    partial class PictureSlidesLabWindow
    {
        private readonly SlideSelectionDialog _gotoSlideDialog = new SlideSelectionDialog();

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

        private void InitGotoSlideDialog()
        {
            _gotoSlideDialog.GetType()
                    .GetProperty("OwningWindow", BindingFlags.Instance | BindingFlags.NonPublic)
                    .SetValue(_gotoSlideDialog, this, null);

            _gotoSlideDialog.OnGotoSlide += () =>
            {
                this.HideMetroDialogAsync(_gotoSlideDialog, MetroDialogOptions);
                if (PowerPointCurrentPresentationInfo.CurrentSlide == null
                    || _gotoSlideDialog.SelectedSlide != PowerPointCurrentPresentationInfo.CurrentSlide.Index)
                {
                    PowerPointPresentation.Current.GotoSlide(_gotoSlideDialog.SelectedSlide);
                    ShowInfoMessageBox(TextCollection.PictureSlidesLabText.SuccessfullyGoToSlide
                        .Replace("_SlideNumber_", _gotoSlideDialog.SelectedSlide.ToString()));
                }
                ViewModel.UpdatePreviewImages(
                    PowerPointCurrentPresentationInfo.CurrentSlide.GetNativeSlide(),
                    PowerPointPresentation.Current.SlideWidth,
                    PowerPointPresentation.Current.SlideHeight);
            };
            _gotoSlideDialog.OnCancel += () =>
            {
                this.HideMetroDialogAsync(_gotoSlideDialog, MetroDialogOptions);
            };
        }
    }
}
