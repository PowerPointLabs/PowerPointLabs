using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;

namespace PowerPointLabs.ImageSearch
{
    partial class ImageSearchPane
    {
        private void InitSearchList()
        {
            SearchList = new ObservableCollection<ImageItem>();
            _downloadedImages = StoragePath.Load(ImagesLabImagesList);
            SearchList.CollectionChanged += SearchList_OnCollectionChanged;

            foreach (var imageItem in _downloadedImages)
            {
                _filesInUse.Add(imageItem.ImageFile);
                _filesInUse.Add(imageItem.FullSizeImageFile);
                if (imageItem.CroppedImageFile != null)
                {
                    _filesInUse.Add(imageItem.CroppedImageFile);
                    _filesInUse.Add(imageItem.CroppedThumbnailImageFile);
                }
            }
            CopyContentToObservableList(_downloadedImages, SearchList);
            Dispatcher.BeginInvoke(new Action(() =>
            {
                SearchListBox.DataContext = this;
            }));
        }

        // rmb to close background presentation
        private void ImageSearchPane_OnClosing(object sender, CancelEventArgs e)
        {
            IsOpen = false;
            IsClosing = true;
            if (PreviewPresentation != null)
            {
                PreviewPresentation.Close();
            }
            if (QuickDropDialog != null)
            {
                QuickDropDialog.Close();
            }
            StoragePath.Save(ImagesLabImagesList, _downloadedImages);
        }

        private void CopyContentToObservableList(IEnumerable<ImageItem> content, ObservableCollection<ImageItem> target)
        {
            foreach (var image in content)
            {
                target.Add(new ImageItem
                {
                    // thumbnail file to show or generate preview images
                    ImageFile = image.ImageFile,
                    // full-size image file
                    FullSizeImageFile = image.FullSizeImageFile,
                    // uri to the full-size image
                    FullSizeImageUri = image.FullSizeImageUri,
                    // link to image's context
                    ContextLink = image.ContextLink,
                    // tooltip to be displayed
                    Tooltip = image.Tooltip,
                    // cropped full-size image file
                    CroppedImageFile = image.CroppedImageFile,
                    // cropped thumbnail file
                    CroppedThumbnailImageFile = image.CroppedThumbnailImageFile,
                    // rectangle used to crop the image
                    Rect = image.Rect
                });
            }
        }
    }
}
