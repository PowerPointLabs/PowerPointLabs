using System;

namespace PowerPointLabs.AutoUpdate.Interface
{
    interface IDownloader
    {
        IDownloader Get(string webAddress, string destinationPath);
        IDownloader After(Downloader.AfterDownloadEventDelegate action);
        IDownloader OnError(Downloader.ErrorEventDelegate action);
        void Start();
    }
}
