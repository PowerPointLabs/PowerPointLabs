using System.Collections.Generic;
using MahApps.Metro.Controls.Dialogs;
using PowerPointLabs.ImagesLab.Model;
using PowerPointLabs.ImagesLab.Thread;
using PowerPointLabs.ImagesLab.Thread.Interface;
using PowerPointLabs.ImagesLab.Util;

namespace PowerPointLabs.ImagesLab.View
{
    public partial class ImagesLabWindow
    {
        ///////////////////////////////////////////////////////////////
        // Implement interface methods
        ///////////////////////////////////////////////////////////////

        public void ShowErrorMessageBox(string content)
        {
            this.ShowMessageAsync("Error", content);
        }

        public void ShowInfoMessageBox(string content)
        {
            this.ShowMessageAsync("Info", content);
        }

        public void ActivateImageDownloadProgressRing()
        {
            ImageDownloadProgressRing.IsActive = true;
        }

        public void DeactivateImageDownloadProgressRing()
        {
            ImageDownloadProgressRing.IsActive = false;
        }

        public void UpdatePreviewImagesForDownloadedImage(ImageItem downloadedImageItem)
        {
            var selectedImageItem = ImageSelectionListBox.SelectedValue as ImageItem;
            if (selectedImageItem != null && downloadedImageItem.ImageFile == selectedImageItem.ImageFile)
            {
                UpdatePreviewImages();
            }
        }

        public IThreadContext GetThreadContext()
        {
            return new ThreadContext(Dispatcher);
        }

        public void ShowSuccessfullyAppliedDialog()
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

        public string InitVariantsComboBox(IDictionary<string, List<StyleVariants>> variants)
        {
            VariantsComboBox.Items.Clear();
            foreach (var key in variants.Keys) { VariantsComboBox.Items.Add(key); }
            VariantsComboBox.SelectedIndex = 0;
            return (string) VariantsComboBox.SelectedValue;
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

        public int GetVariationListBoxSelectedId()
        {
            return StylesVariationListBox.SelectedIndex >= 0 ? StylesVariationListBox.SelectedIndex : 0;
        }

        public void SetVariationListBoxSelectedId(int index)
        {
            StylesVariationListBox.SelectedIndex = index;
        }

        public string GetVariantsComboBoxSelectedValue()
        {
            return (string) VariantsComboBox.SelectedValue;
        }
    }
}
