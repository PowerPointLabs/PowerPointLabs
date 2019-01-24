using System;
using System.Reflection;
using System.Windows;

using MahApps.Metro.Controls.Dialogs;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.PictureSlidesLab.Views
{
    partial class PictureSlidesLabWindow
    {
        private readonly SlideSelectionDialog _gotoSlideDialog = new SlideSelectionDialog();

        private void GotoSlideButton_OnClick(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_gotoSlideDialog.IsOpen)
                {
                    return;
                }

                _gotoSlideDialog
                    .Init("Go to a Slide to Edit")
                    .CustomizeGotoSlideButton("Go", "Go to the slide to edit its style.")
                    .FocusOkButton()
                    .OpenDialog();
                this.ShowMetroDialogAsync(_gotoSlideDialog, MetroDialogOptions);
            }
            catch (Exception expt)
            {
                Logger.LogException(expt, "GotoSlideButton_OnClick");
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
            Logger.Log("PSL init GotoSlideDialog done");
        }

        private void GotoSlideWithStyleLoading()
        {
            _gotoSlideDialog.CloseDialog();
            this.HideMetroDialogAsync(_gotoSlideDialog, MetroDialogOptions);

            GotoSlide();

            LoadStyleAndImage(this.GetCurrentPresentation()
                .Slides[_gotoSlideDialog.SelectedSlide - 1]);
        }

        private void GotoSlide()
        {
            if (this.GetCurrentSlide() == null
                || _gotoSlideDialog.SelectedSlide != this.GetCurrentSlide().Index)
            {
                this.GetCurrentPresentation().GotoSlide(_gotoSlideDialog.SelectedSlide);
            }
            UpdatePreviewImages();
        }
    }
}
