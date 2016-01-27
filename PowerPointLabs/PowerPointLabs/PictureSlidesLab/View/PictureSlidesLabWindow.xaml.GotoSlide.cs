using System.Reflection;
using System.Windows;
using MahApps.Metro.Controls.Dialogs;
using PowerPointLabs.Models;

namespace PowerPointLabs.PictureSlidesLab.View
{
    partial class PictureSlidesLabWindow
    {
        private readonly SlideSelectionDialog _gotoSlideDialog = new SlideSelectionDialog();
        private bool _isDisplayDefaultPicture;

        private void GotoSlideButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_gotoSlideDialog.IsOpen) return;

                _gotoSlideDialog
                    .Init("Select the Slide to Edit")
                    .CustomizeGotoSlideButton("Select", "Select the slide to edit styles.")
                    .FocusOkButton()
                    .OpenDialog();
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

            _gotoSlideDialog.OnGotoSlide += GotoSlideWithStyleLoading;

            _gotoSlideDialog.OnCancel += () =>
            {
                _gotoSlideDialog.CloseDialog();
                this.HideMetroDialogAsync(_gotoSlideDialog, MetroDialogOptions);
            };
        }

        private void GotoSlideWithStyleLoading()
        {
            this.HideMetroDialogAsync(_gotoSlideDialog, MetroDialogOptions);

            GotoSlide();

            LoadStyleAndImage(PowerPointPresentation.Current
                .Slides[_gotoSlideDialog.SelectedSlide - 1]);
        }

        private void GotoSlide()
        {
            if (PowerPointCurrentPresentationInfo.CurrentSlide == null
                || _gotoSlideDialog.SelectedSlide != PowerPointCurrentPresentationInfo.CurrentSlide.Index)
            {
                PowerPointPresentation.Current.GotoSlide(_gotoSlideDialog.SelectedSlide);
            }
            UpdatePreviewImages();
        }
    }
}
