namespace PowerPointLabs.AutoUpdate.Interface
{
    public interface IDownloader
    {
        IDownloader Get(string webAddress, string destinationPath);
        IDownloader After(Downloader.AfterDownloadEventDelegate action);
        IDownloader OnError(Downloader.ErrorEventDelegate action);
        void Start();
    }
}
