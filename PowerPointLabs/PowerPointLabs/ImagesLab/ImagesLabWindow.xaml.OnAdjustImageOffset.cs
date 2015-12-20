using System.Windows;
using PowerPointLabs.ImagesLab.Domain;
using PowerPointLabs.ImagesLab.Util;

namespace PowerPointLabs.ImagesLab
{
    partial class ImagesLabWindow
    {
        public AdjustImageWindow CropWindow { get; set; }

        private void AdjustImageOffset(ImageItem imageItem = null)
        {
            if (imageItem == null && ImageSelectionListBox.SelectedValue == null) return;

            var source = imageItem ?? (ImageItem)ImageSelectionListBox.SelectedValue;
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

            if (_rightClickedSearchListBoxItemIndex < 0
                || _rightClickedSearchListBoxItemIndex > ImageSelectionListBox.Items.Count)
                return;

            var selectedImage = (ImageItem)ImageSelectionListBox.Items.GetItemAt(_rightClickedSearchListBoxItemIndex);
            if (selectedImage == null) return;

            AdjustImageOffset(selectedImage);
        }

        private void MenuItemAdjustImage_OnClickFromListBox(object sender, RoutedEventArgs e)
        {
            var selectedImage = (ImageItem)ImageSelectionListBox.SelectedItem;
            if (selectedImage == null || selectedImage.ImageFile == StoragePath.LoadingImgPath) return;

            AdjustImageOffset(selectedImage);
        }
    }
}
