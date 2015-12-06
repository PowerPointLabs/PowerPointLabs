using System;
using System.Collections.Generic;
using MahApps.Metro.Controls.Dialogs;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;

namespace PowerPointLabs.ImageSearch
{
    public partial class ImageSearchPane
    {
        ///////////////////////////////////////////////////////////////
        /// Common
        ///////////////////////////////////////////////////////////////

        private void HandleDownloadedThumbnail(
            ImageItem item, string thumbnailPath, object searchResult = null, IList<int> selectedIds = null)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (item == null) return;
                item.ImageFile = thumbnailPath;

                if (searchResult != null)
                {
                    item.FullSizeImageUri = VOUtil.GetLink(searchResult);
                    item.Tooltip = GetTooltip(searchResult);
                    item.ContextLink = VOUtil.GetContextLink(searchResult);
                }
                else // use case download image & when thumbnail is already full-size
                {
                    item.FullSizeImageFile = item.ImageFile;
                }

                var selectedImageItem = SearchListBox.SelectedValue as ImageItem;
                if (selectedImageItem != null && item.ImageFile == selectedImageItem.ImageFile)
                {
                    DoPreview(selectedIds);
                }
            }));
        }

        // timer's downloading will come here at the end,
        // or both timer + insert's downloading will come here
        private void HandleDownloadedFullSizeImage(ImageItem source, string fullsizeImageFile)
        {
            // in downloader thread
            Dispatcher.BeginInvoke(new Action(() =>
            {
                if (source == null) return;

                // UI thread again
                // store back to image, so cache it
                source.FullSizeImageFile = fullsizeImageFile;
                var fullsizeImageUri = source.FullSizeImageUri;

                // intent: during download, selected item may have been changed to another one
                // if selected one got changed,
                // 1. no need to preview it
                // 2. no need to insert it to current slide
                var currentImageItem = SearchListBox.SelectedValue as ImageItem;
                if (currentImageItem == null)
                {
                    PreviewProgressRing.IsActive = false;
                }
                else if (currentImageItem.ImageFile == source.ImageFile)
                {
                    // if selected one remains
                    // and it is to insert the full size image,
                    if (_applyDownloadingUriList.Contains(fullsizeImageUri))
                    {
                        // open confirm apply flyout + do preview
                        if (PreviewListBox.SelectedValue != null)
                        {
                            OpenConfirmApplyFlyout(PreviewListBox.SelectedItems);
                        }
                        DoPreview(source);
                    }
                    // or it is to preview only (from timer)
                    else if (_timerDownloadingUriList.Contains(fullsizeImageUri))
                    {
                        DoPreview(source);
                    }
                }

                RemoveDebounceCheck(fullsizeImageUri);
            }));
        }

        private void RemoveDebounceCheck(string fullsizeImageUri)
        {
            _applyDownloadingUriList.Remove(fullsizeImageUri);
            _timerDownloadingUriList.Remove(fullsizeImageUri);
        }

        private void ShowErrorMessageBox(string content)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                this.ShowMessageAsync("Error", content);
            }));
        }
    }
}
