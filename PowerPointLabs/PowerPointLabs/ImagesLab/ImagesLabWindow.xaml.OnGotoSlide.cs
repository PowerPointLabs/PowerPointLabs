using System;
using System.Reflection;
using System.Windows;
using MahApps.Metro.Controls.Dialogs;
using PowerPointLabs.Models;

namespace PowerPointLabs.ImagesLab
{
    partial class ImagesLabWindow
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
                if (_gotoSlideDialog.SelectedSlide > 0)
                {
                    PowerPointPresentation.Current.GotoSlide(_gotoSlideDialog.SelectedSlide);
                }
                _latestImageChangedTime = DateTime.Now;
                DoPreview();
            };
            _gotoSlideDialog.OnCancel += () =>
            {
                this.HideMetroDialogAsync(_gotoSlideDialog, MetroDialogOptions);
            };
        }
    }
}
