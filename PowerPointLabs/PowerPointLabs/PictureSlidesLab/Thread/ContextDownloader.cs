using PowerPointLabs.AutoUpdate;
using PowerPointLabs.AutoUpdate.Interface;
using PowerPointLabs.PictureSlidesLab.Thread.Interface;

namespace PowerPointLabs.PictureSlidesLab.Thread
{
    internal class ContextDownloader : IDownloader
    {
        private string _downloadLink;
        private string _destination;
        private Downloader.AfterDownloadEventDelegate _onAfterDownload;
        private Downloader.ErrorEventDelegate _onError;
        private IThreadContext _threadContext;
        private IDownloader _downloader;

        public ContextDownloader(IThreadContext threadContext)
        {
            _threadContext = threadContext;
        }

        public IDownloader SetDownloader(IDownloader downloader)
        {
            _downloader = downloader;
            return this;
        }

        public IDownloader Get(string webAddress, string destinationPath)
        {
            _downloadLink = webAddress;
            _destination = destinationPath;
            return this;
        }

        public IDownloader After(Downloader.AfterDownloadEventDelegate action)
        {
            _onAfterDownload = () =>
            {
                _threadContext.BeginInvoke(() =>
                {
                    action();
                });
            };
            return this;
        }

        public IDownloader OnError(Downloader.ErrorEventDelegate action)
        {
            _onError = e =>
            {
                _threadContext.BeginInvoke(() =>
                {
                    action(e);
                });
            };
            return this;
        }

        public void Start()
        {
            IDownloader downloader = _downloader ?? new Downloader();
            downloader
                .Get(_downloadLink, _destination)
                .After(_onAfterDownload)
                .OnError(_onError)
                .Start();
        }
    }
}
