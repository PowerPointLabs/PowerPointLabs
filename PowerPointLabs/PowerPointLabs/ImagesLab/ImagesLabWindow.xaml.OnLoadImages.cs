using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPointLabs.AutoUpdate;
using PowerPointLabs.ImagesLab.Domain;
using PowerPointLabs.ImagesLab.Util;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ImagesLab
{
    public partial class ImagesLabWindow
    {
        # region Internal APIs

        private void DoLoadImageFromFile(string[] filenames = null)
        {
            Dispatcher.Invoke(new Action(() =>
            {
                if (filenames == null)
                {
                    var openFileDialog = new OpenFileDialog
                    {
                        Multiselect = true,
                        Filter = @"Image File|*.png;*.jpg;*.jpeg;*.bmp;*.gif;"
                    };
                    var fileDialogResult = openFileDialog.ShowDialog();
                    if (fileDialogResult != System.Windows.Forms.DialogResult.OK)
                    {
                        return;
                    }
                    filenames = openFileDialog.FileNames;
                }

                try
                {
                    foreach (var filename in filenames)
                    {
                        VerifyIsProperImage(filename);
                        var fromFileItem = new ImageItem
                        {
                            ImageFile = ImageUtil.GetThumbnailFromFullSizeImg(filename),
                            FullSizeImageFile = filename,
                            FullSizeImageUri = filename,
                            ContextLink = filename,
                            Tooltip = ImageUtil.GetWidthAndHeight(filename)
                        };
                        //add it
                        ImageSelectionList.Add(fromFileItem);  
                    }
                }
                catch
                {
                    // not an image or image is corrupted
                    ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorImageCorrupted);
                }
            }));
        }

        private static void VerifyIsProperImage(string filename)
        {
            using (Image.FromFile(filename))
            {
                // so this is a proper image
            }
        }

        private void DoDownloadImage(string downloadLink = null)
        {
            Dispatcher.Invoke(new Action(() =>
            {
                if (StringUtil.IsEmpty(downloadLink))
                {
                    return;
                }
                // Error Case 1: If url not valid
                if (!UrlUtil.IsUrlValid(downloadLink))
                {
                    ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorUrlLinkIncorrect);
                    return;
                }

                var item = new ImageItem
                {
                    ImageFile = StoragePath.LoadingImgPath,
                    ContextLink = downloadLink
                };
                UrlUtil.GetMetaInfo(ref downloadLink, item);

                // add it
                ImageSelectionList.Add(item);
                ImageDownloadProgressRing.IsActive = true;

                var imagePath = StoragePath.GetPath("img-" 
                    + DateTime.Now.GetHashCode() + "-" 
                    + Guid.NewGuid().ToString().Substring(0, 7));
                new Downloader()
                    .Get(downloadLink, imagePath)
                    .After(() =>
                    {
                        try
                        {
                            // Error Case 2: not a proper image
                            VerifyIsProperImage(imagePath);

                            Dispatcher.Invoke(new Action(() =>
                            {
                                // TODO turn off progress ring after all downloaded
                                ImageDownloadProgressRing.IsActive = false;
                            }));
                            HandleDownloadedThumbnail(item, imagePath);
                        }
                        catch
                        {
                            ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorImageDownloadCorrupted);
                            Dispatcher.Invoke(new Action(() =>
                            {
                                // TODO turn off progress ring after all downloaded
                                ImageDownloadProgressRing.IsActive = false;
                                ImageSelectionList.Remove(item);
                            }));
                        }
                    })
                    // Error Case 3: Possibly network timeout
                    .OnError(e => { RemoveImageItem(item, e); })
                    .Start();
            }));
        }

        # endregion

        # region Helper Funcs

        private void RemoveImageItem(ImageItem item, Exception e)
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                ImageDownloadProgressRing.IsActive = false;
                ImageSelectionList.Remove(item);
            }));
            ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorFailedToLoad + e.Message);
        }
        # endregion
    }
}
