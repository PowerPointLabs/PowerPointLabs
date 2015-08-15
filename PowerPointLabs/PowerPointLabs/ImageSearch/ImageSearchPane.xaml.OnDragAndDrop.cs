using System.Windows;
using PowerPointLabs.WPF.Observable;

namespace PowerPointLabs.ImageSearch
{
    public partial class ImageSearchPane
    {
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

        private void OnDragLeave(object sender, DragEventArgs args)
        {
            ImagesLabGridOverlay.Visibility = Visibility.Hidden;
        }

        private void OnDragEnter(object sender, DragEventArgs args)
        {
            if (args.Data.GetDataPresent("FileDrop")
                || args.Data.GetDataPresent("Text"))
            {
                CloseFlyouts();
                ImagesLabGridOverlay.Visibility = Visibility.Visible;
                ActivateWithoutPreview();
            }
        }

        private void CloseFlyouts()
        {
            if (StyleOptionsFlyout.IsOpen)
            {
                StyleOptionsFlyout.IsOpen = false;
            }
            if (SearchOptionsFlyout.IsOpen)
            {
                SearchOptionsFlyout.IsOpen = false;
            }
            if (ConfirmApplyFlyout.IsOpen)
            {
                ConfirmApplyFlyout.IsOpen = false;
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

                    var filename = filenames[0];
                    DoLoadImageFromFile(filename);
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
    }
}
