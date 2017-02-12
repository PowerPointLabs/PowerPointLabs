﻿using System;
using System.Net;
using System.Threading.Tasks;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.AutoUpdate.Interface;

namespace PowerPointLabs.AutoUpdate
{
    public class Downloader : IDownloader
    {
        private readonly WebClient _client = new WebClient();

        public delegate void AfterDownloadEventDelegate();
        private event AfterDownloadEventDelegate AfterDownload;

        public delegate void ErrorEventDelegate(Exception e);
        private event ErrorEventDelegate WhenError;

        private string _downloadAddress = "";
        private string _destAddress = "";

        public Downloader()
        {
            //cancel default proxy, which may use IE's proxy settings
            _client.Proxy = null;
        }

        public IDownloader Get(string webAddress, string destinationPath)
        {
            _downloadAddress = webAddress;
            _destAddress = destinationPath;
            return this;
        }

        public IDownloader After(AfterDownloadEventDelegate action)
        {
            AfterDownload = action;
            return this;
        }

        public IDownloader OnError(ErrorEventDelegate action)
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
                Logger.LogException(e, "Failed to start thread of Downloader.StartDownload");
            }
        }

        private void CallAfterDownloadDelegate()
        {
            var handler = AfterDownload;
            if (handler != null) handler();
        }

        private void CallWhenErrorDelegate(Exception e)
        {
            var handler = WhenError;
            if (handler != null) handler(e);
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
                CallWhenErrorDelegate(e);
                Logger.LogException(e, "Failed to execute Downloader.StartDownload");
            }
        }
    }
}
