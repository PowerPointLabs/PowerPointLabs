using System.Windows;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;

namespace PowerPointLabs.ImageSearch
{
    partial class ImageSearchPane
    {
        public AdjustImageWindow CropWindow { get; set; }

        private void AdjustImageOffset(ImageItem imageItem = null)
        {
            if (imageItem == null && SearchListBox.SelectedValue == null) return;

            var source = imageItem ?? (ImageItem)SearchListBox.SelectedValue;
            CropWindow = new AdjustImageWindow();
            CropWindow.SetThumbnailImage(source.ImageFile);
            CropWindow.SetFullsizeImage(source.FullSizeImageFile);
            if (source.Rect.Width > 1)
            {
                CropWindow.SetCropRect(source.Rect.X, source.Rect.Y, source.Rect.Width, source.Rect.Height);
            }
            CropWindow.IsOpen = true;
            CropWindow.ShowDialog();
            CropWindow.IsOpen = false;

            if (CropWindow.IsCropped)
            {
                source.CroppedImageFile = CropWindow.CropResult;
                source.CroppedThumbnailImageFile = CropWindow.CropResultThumbnail;
                source.Rect = CropWindow.Rect;
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

        private void MenuItemAdjustImage_OnClickFromListBox(object sender, RoutedEventArgs e)
        {
            var selectedImage = (ImageItem)SearchListBox.SelectedItem;
            if (selectedImage == null || selectedImage.ImageFile == StoragePath.LoadingImgPath) return;

            AdjustImageOffset(selectedImage);
        }
    }
}
