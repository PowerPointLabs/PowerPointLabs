using System;
using System.Timers;
using PowerPointLabs.AutoUpdate;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.Util;

namespace PowerPointLabs.ImageSearch
{
    // TODO: this partial class is obsolete
    public partial class ImageSearchPane
    {
        // intent:
        // when select a thumbnail for some time (defined by TimerInterval),
        // try to download its full size version for better preview and can be used for insertion
        private void InitPreviewTimer()
        {
            PreviewTimer = new Timer { Interval = TimerInterval };
            PreviewTimer.Elapsed += (sender, args) =>
            {
                // in timer thread
                PreviewTimer.Stop();
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    // UI thread starts
                    var source = SearchListBox.SelectedValue as ImageItem;
                    // if already have cached full-size image, ignore
                    if (source == null || source.FullSizeImageFile != null)
                    {
                        // do nothing
                    }
                    // if not downloading the full size image yet, download it
                    else if (!_timerDownloadingUriList.Contains(source.FullSizeImageUri))
                    {
                        _timerDownloadingUriList.Add(source.FullSizeImageUri);
                        // preview progress ring will be off, after preview processing is done
                        SetProgressingRingStatus(true);

                        var fullsizeImageFile = TempPath.GetPath("fullsize");
                        new Downloader()
                            .Get(source.FullSizeImageUri, fullsizeImageFile)
                            .After(()=> { HandleDownloadedFullSizeImage(source, fullsizeImageFile); })
                            .OnError(WhenFailDownloadFullSizeImage)
                            .Start();
                    }
                    // it's downloading
                    else
                    {
                        // preview progress ring will be off, after preview processing is done
                        SetProgressingRingStatus(true);
                    }
                }));
            };
        }

        private void WhenFailDownloadFullSizeImage()
        {
            Dispatcher.BeginInvoke(new Action(() =>
            {
                SetProgressingRingStatus(false);
            }));
        }
    }
}
