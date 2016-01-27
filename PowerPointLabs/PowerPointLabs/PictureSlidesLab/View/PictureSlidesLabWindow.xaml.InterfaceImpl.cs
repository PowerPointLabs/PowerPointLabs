using System.Windows.Media;
using MahApps.Metro.Controls.Dialogs;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.PictureSlidesLab.Thread;
using PowerPointLabs.PictureSlidesLab.Thread.Interface;
using PowerPointLabs.PictureSlidesLab.Util;

namespace PowerPointLabs.PictureSlidesLab.View
{
    public partial class PictureSlidesLabWindow
    {
        ///////////////////////////////////////////////////////////////
        // Implemented interface methods
        ///////////////////////////////////////////////////////////////

        public void ShowErrorMessageBox(string content)
        {
            this.ShowMessageAsync("Error", content);
        }

        public void ShowInfoMessageBox(string content)
        {
            this.ShowMessageAsync("Info", content);
        }

        public void ShowSuccessfullyAppliedDialog()
        {
            try
            {
                if (_gotoSlideDialog.IsOpen) return;

                _gotoSlideDialog
                    .Init("Successfully Applied!")
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

        public IThreadContext GetThreadContext()
        {
            return new ThreadContext(Dispatcher);
        }

        public double GetVariationListBoxScrollOffset()
        {
            var scrollOffset = 0d;
            var scrollViewer = ListBoxUtil.FindScrollViewer(StylesVariationListBox);
            if (scrollViewer != null) { scrollOffset = scrollViewer.VerticalOffset; }
            return scrollOffset;
        }

        public void SetVariationListBoxScrollOffset(double offset)
        {
            var scrollViewer = ListBoxUtil.FindScrollViewer(StylesVariationListBox);
            if (scrollViewer != null) { scrollViewer.ScrollToVerticalOffset(offset); }
        }

        public void SetVariantsColorPanelBackground(Brush color)
        {
            VariantsColorPanel.Background = color;
        }

        public ImageItem CreateDefaultPictureItem()
        {
            return new ImageItem
            {
                ImageFile = StoragePath.NoPicturePlaceholderImgPath,
                Tooltip = "Please select a picture."
            };
        }

        public bool IsDisplayDefaultPicture()
        {
            return _isDisplayDefaultPicture;
        }

        public void EnableUpdatingPreviewImages()
        {
            _isDisplayDefaultPicture = false;
        }

        public void DisableUpdatingPreviewImages()
        {
            _isDisplayDefaultPicture = true;
        }
    }
}
