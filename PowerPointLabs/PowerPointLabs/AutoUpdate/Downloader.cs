using System;
using System.Net;
using System.Threading;

namespace PowerPointLabs.AutoUpdate
{
    class Downloader
    {
        private readonly WebClient _client = new WebClient();

        public delegate void AfterDownloadEventDelegate();
        public event AfterDownloadEventDelegate AfterDownload;

        private String _downloadAddress = "";
        private String _destAddress = "";

        public Downloader()
        {
            //cancel default proxy, which may use IE's proxy settings
            _client.Proxy = null;
        }

        private void OnAfterDownload()
        {
            var handler = AfterDownload;
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

        public void Start()
        {
            var th = new Thread(new ThreadStart(StartDownload));
            th.Start();
        }

        private void StartDownload()
        {
            if (_downloadAddress == "" || _destAddress == "") 
                return;

            _client.DownloadFile(_downloadAddress, _destAddress);
            OnAfterDownload();
        }
    }
}
