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
            var loadedImages = StoragePath.Load(ImagesLabImagesList);
            SearchList.CollectionChanged += SearchList_OnCollectionChanged;

            foreach (var imageItem in loadedImages)
            {
                _filesInUse.Add(imageItem.ImageFile);
                _filesInUse.Add(imageItem.FullSizeImageFile);
                if (imageItem.CroppedImageFile != null)
                {
                    _filesInUse.Add(imageItem.CroppedImageFile);
                    _filesInUse.Add(imageItem.CroppedThumbnailImageFile);
                }
            }
            CopyContentToObservableList(loadedImages, SearchList);
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
            StoragePath.Save(ImagesLabImagesList, SearchList);
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
