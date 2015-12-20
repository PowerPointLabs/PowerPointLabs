using System;
using MahApps.Metro.Controls.Dialogs;
using PowerPointLabs.ImagesLab.Domain;
using PowerPointLabs.ImagesLab.Util;

namespace PowerPointLabs.ImagesLab
{
    public partial class ImagesLabWindow
    {
        ///////////////////////////////////////////////////////////////
        /// Common
        ///////////////////////////////////////////////////////////////

        private void SetProgressingRingStatus(bool isActive)
        {
            PreviewProgressRing.IsActive = isActive;
            VariationProgressRing.IsActive = isActive;
        }

        private void HandleDownloadedThumbnail(
            ImageItem item, string thumbnailPath)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (item == null) return;
                item.ImageFile = thumbnailPath;
                item.FullSizeImageFile = item.ImageFile;
                item.ImageFile = ImageUtil.GetThumbnailFromFullSizeImg(item.FullSizeImageFile);
                item.Tooltip = ImageUtil.GetWidthAndHeight(item.FullSizeImageFile);

                var selectedImageItem = ImageSelectionListBox.SelectedValue as ImageItem;
                if (selectedImageItem != null && item.ImageFile == selectedImageItem.ImageFile)
                {
                    DoPreview();
                }
            }));
        }

        private void ShowErrorMessageBox(string content)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                this.ShowMessageAsync("Error", content);
            }));
        }

        private void ShowInfoMessageBox(string content)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                this.ShowMessageAsync("Info", content);
            }));
        }

        private void OpenSuccessfullyAppliedDialog()
        {
            try
            {
                _gotoSlideDialog.Init("Successfully Applied!");
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
