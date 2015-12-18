using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using PowerPointLabs.ImagesLab.Domain;
using PowerPointLabs.ImagesLab.Util;

namespace PowerPointLabs.ImagesLab
{
    partial class ImagesLabWindow
    {
        private const string ImagesLabImagesList = "ImagesLabImagesList";

        private void InitImageSelectionList()
        {
            ImageSelectionList = new ObservableCollection<ImageItem>();
            var loadedImages = StoragePath.Load(ImagesLabImagesList);
            ImageSelectionList.CollectionChanged += ImageSelectionList_OnCollectionChanged;

            foreach (var imageItem in loadedImages)
            {
                _imageFilesInUse.Add(imageItem.ImageFile);
                _imageFilesInUse.Add(imageItem.FullSizeImageFile);
                if (imageItem.CroppedImageFile != null)
                {
                    _imageFilesInUse.Add(imageItem.CroppedImageFile);
                    _imageFilesInUse.Add(imageItem.CroppedThumbnailImageFile);
                }
            }
            CopyContentToObservableList(loadedImages, ImageSelectionList);
            Dispatcher.BeginInvoke(new Action(() =>
            {
                ImageSelectionListBox.DataContext = this;
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
            StoragePath.Save(ImagesLabImagesList, ImageSelectionList);
        }

        private void CopyContentToObservableList(IEnumerable<ImageItem> content, ObservableCollection<ImageItem> target)
        {
            foreach (var image in content)
            {
                target.Add(image);
            }
        }
    }
}
