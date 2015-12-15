using System;
using System.Windows;
using PowerPointLabs.WPF.Observable;

namespace PowerPointLabs.ImageSearch
{
    public partial class ImageSearchPane
    {
        public QuickDropDialog QuickDropDialog { get; set; }

        private void InitDragAndDrop()
        {
            DragAndDropInstructionText = new ObservableString { Text = "Drag and Drop here to get image." };
            DragAndDropInstructions.DataContext = DragAndDropInstructionText;

            AllowDrop = true;

            DragEnter += OnDragEnter;
            DragLeave += OnDragLeave;
            DragOver += OnDragEnter;
            Drop += OnDrop;
        }

        private void InitQuickDropDialog()
        {
            QuickDropDialog = new QuickDropDialog(this);
            QuickDropDialog.DropHandler += OnDrop;
        }

        private void OnDragLeave(object sender, DragEventArgs args)
        {
            ImagesLabGridOverlay.Visibility = Visibility.Hidden;
        }

        private void OnDragEnter(object sender, DragEventArgs args)
        {
            if (args.Data.GetDataPresent("FileDrop")
                || args.Data.GetDataPresent("Text"))
            {
                ImagesLabGridOverlay.Visibility = Visibility.Visible;
                ActivateWithoutPreview();
            }
        }

        private void CloseFlyouts()
        {
            if (_isVariationsFlyoutOpen)
            {
                CloseVariationsFlyout();
            }
            if (SearchOptionsFlyout.IsOpen)
            {
                SearchOptionsFlyout.IsOpen = false;
            }
            if (_isCustomizationFlyoutOpen)
            {
                CloseCustomizationFlyout();
            }
        }

        private void ActivateWithoutPreview()
        {
            _isWindowActivatedWithPreview = false;
            Activate();
            _isWindowActivatedWithPreview = true;
        }

        private void OnDrop(object sender, DragEventArgs args)
        {
            try
            {
                if (args == null) return;

                if (args.Data.GetDataPresent("FileDrop"))
                {
                    var filenames = (args.Data.GetData("FileDrop") as string[]);
                    if (filenames == null || filenames.Length == 0) return;

                    DoLoadImageFromFile(filenames);
                }
                else if (args.Data.GetDataPresent("Text"))
                {
                    var imageUrl = args.Data.GetData("Text") as string;
                    DoDownloadImage(imageUrl);
                }
            }
            finally
            {
                ImagesLabGridOverlay.Visibility = Visibility.Hidden;
            }
        }

        private void ImageSearchPane_OnDeactivated(object sender, EventArgs e)
        {
            if (!IsClosing 
                && (CropWindow == null || !CropWindow.IsOpen) 
                && (QuickDropDialog == null || !QuickDropDialog.IsOpen))
            {
                InitQuickDropDialog();
                QuickDropDialog.Show();
            }
        }
    }
}
