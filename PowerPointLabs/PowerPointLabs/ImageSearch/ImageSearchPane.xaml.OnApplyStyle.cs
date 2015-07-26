using PowerPointLabs.AutoUpdate;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;

namespace PowerPointLabs.ImageSearch
{
    public partial class ImageSearchPane
    {
        private void ApplyStyle()
        {
            PreviewTimer.Stop();
            PreviewProgressRing.IsActive = true;

            var source = (ImageItem)SearchListBox.SelectedValue;
            var targetStyle = (ImageItem)PreviewListBox.SelectedValue;
            if (source == null || targetStyle == null) return;

            if (source.FullSizeImageFile != null)
            {
                PreviewPresentation.ApplyStyle(source, targetStyle.Tooltip);
                PreviewProgressRing.IsActive = false;
            }
            else if (!_insertDownloadingUriList.Contains(source.FullSizeImageUri))
            {
                var fullsizeImageUri = source.FullSizeImageUri;
                _insertDownloadingUriList.Add(fullsizeImageUri);
                _insertDownloadingUriToPreviewImage[fullsizeImageUri] = targetStyle;

                var fullsizeImageFile = TempPath.GetPath("fullsize");
                new Downloader()
                    .Get(fullsizeImageUri, fullsizeImageFile)
                    .After(() => { HandleDownloadedFullSizeImage(source, fullsizeImageFile); })
                    .OnError(() => { ShowErrorMessageBox(TextCollection.ImagesLabText.ErrorNetworkOrSourceUnavailable); })
                    .Start();
            }
            // already downloading, then update preview image in the map
            else
            {
                var fullsizeImageUri = source.FullSizeImageUri;
                _insertDownloadingUriToPreviewImage[fullsizeImageUri] = targetStyle;
            }
        }
    }
}
