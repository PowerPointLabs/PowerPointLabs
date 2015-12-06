using System;
using System.Net;
using System.Threading.Tasks;

namespace PowerPointLabs.AutoUpdate
{
    class Downloader
    {
        private readonly WebClient _client = new WebClient();

        public delegate void AfterDownloadEventDelegate();
        private event AfterDownloadEventDelegate AfterDownload;

        public delegate void ErrorEventDelegate();
        private event ErrorEventDelegate WhenError;

        private String _downloadAddress = "";
        private String _destAddress = "";

        public Downloader()
        {
            //cancel default proxy, which may use IE's proxy settings
            _client.Proxy = null;
        }

        private void CallAfterDownloadDelegate()
        {
            var handler = AfterDownload;
            if (handler != null) handler();
        }

        private void CallWhenErrorDelegate()
        {
            var handler = WhenError;
            if (handler != null) handler();
        }

        public Downloader Get(String webAddress, String destinationPath)
        {
            _downloadAddress = webAddress;
            _destAddress = destinationPath;
            return this;
        }

        public Downloader After(AfterDownloadEventDelegate action)
        {
            AfterDownload = action;
            return this;
        }

        public Downloader OnError(ErrorEventDelegate action)
        {
            WhenError = action;
            return this;
        }

        public void Start()
        {
            try
            {
                var th = new Task(StartDownload);
                th.Start();
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.LogException(e, "Failed to start thread of Downloader.StartDownload");
            }
        }

        private void StartDownload()
        {
            if (_downloadAddress == "" || _destAddress == "") 
                return;

            try
            {
                _client.DownloadFile(_downloadAddress, _destAddress);
                CallAfterDownloadDelegate();
            }
            catch (Exception e)
            {
                CallWhenErrorDelegate();
                PowerPointLabsGlobals.LogException(e, "Failed to execute Downloader.StartDownload");
            }
        }
    }
}
