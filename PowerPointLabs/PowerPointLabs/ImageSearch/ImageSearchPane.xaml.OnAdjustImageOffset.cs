using System.Windows;
using PowerPointLabs.ImageSearch.Crop;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;

namespace PowerPointLabs.ImageSearch
{
    partial class ImageSearchPane
    {
        private void AdjustImageOffset(ImageItem imageItem = null)
        {
            if (imageItem == null && SearchListBox.SelectedValue == null) return;

            var source = imageItem ?? (ImageItem)SearchListBox.SelectedValue;
            var cropWindow = new CropWindow();
            cropWindow.SetThumbnailImage(source.ImageFile);
            cropWindow.SetFullsizeImage(source.FullSizeImageFile);
            if (source.Rect.Width > 1)
            {
                cropWindow.SetCropRect(source.Rect.X, source.Rect.Y, source.Rect.Width, source.Rect.Height);
            }
            cropWindow.ShowDialog();

            if (cropWindow.IsCropped)
            {
                source.CroppedImageFile = cropWindow.CropResult;
                source.CroppedThumbnailImageFile = cropWindow.CropResultThumbnail;
                source.Rect = cropWindow.Rect;

                var imageIndex = SearchListBox.Items.IndexOf(source);
                if (imageIndex >= 0
                    && imageIndex < _downloadedImages.Count)
                {
                    var imageToPersist = _downloadedImages[imageIndex];
                    imageToPersist.CroppedImageFile = source.CroppedImageFile;
                    imageToPersist.CroppedThumbnailImageFile = source.CroppedThumbnailImageFile;
                    imageToPersist.Rect = source.Rect;
                }
            }
        }

        private void MenuItemAdjustImage_OnClick(object sender, RoutedEventArgs e)
        {
            if (SearchButton.SelectedIndex == TextCollection.ImagesLabText.ButtonIndexSearch) return;

            if (_rightClickedSearchListBoxItemIndex < 0
                || _rightClickedSearchListBoxItemIndex > SearchListBox.Items.Count)
                return;

            var selectedImage = (ImageItem)SearchListBox.Items.GetItemAt(_rightClickedSearchListBoxItemIndex);
            if (selectedImage == null) return;

            AdjustImageOffset(selectedImage);
        }

        private void MenuItemAdjustImage_OnClickFromPreviewListBox(object sender, RoutedEventArgs e)
        {
            var selectedImage = (ImageItem)SearchListBox.SelectedItem;
            if (selectedImage == null || selectedImage.ImageFile == StoragePath.LoadingImgPath) return;

            AdjustImageOffset(selectedImage);
        }
    }
}
