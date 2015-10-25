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
            if (PreviewPresentation != null)
            {
                PreviewPresentation.Close();
            }
            StoragePath.Save(ImagesLabImagesList, _downloadedImages);
        }

        private void CopyContentToObservableList(IEnumerable<ImageItem> content, ObservableCollection<ImageItem> target)
        {
            foreach (var image in content)
            {
                target.Add(new ImageItem
                {
                    ImageFile = image.ImageFile,
                    FullSizeImageFile = image.FullSizeImageFile,
                    FullSizeImageUri = image.FullSizeImageUri,
                    ContextLink = image.ContextLink,
                    Tooltip = image.Tooltip,
                    CroppedImageFile = image.CroppedImageFile,
                    CroppedThumbnailImageFile = image.CroppedThumbnailImageFile,
                    Rect = image.Rect
                });
            }
        }
    }
}
