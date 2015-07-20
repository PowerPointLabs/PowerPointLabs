using System;
using System.Net;
using System.Threading;
using PowerPointLabs.ImageSearch.SearchEngine.Options;
using PowerPointLabs.ImageSearch.SearchEngine.VO;
using RestSharp;
using RestSharp.Deserializers;

namespace PowerPointLabs.ImageSearch.SearchEngine
{
    public class GoogleEngine
    {
        public const int NumOfItemsPerSearch = 30;
        public const int NumOfItemsPerRequest = 10;

        // state, used for Search More
        private string _lastTimeQuery = "";
        private int _nextStartIndex;

        // multi thread state
        private readonly Object _syncLock = new object();
        private bool _isFailedAlready;

        public GoogleOptions GoogleOptions { get; set; }

        public GoogleEngine(GoogleOptions options)
        {
            GoogleOptions = options;
        }

        public delegate void WhenFailEventDelegate(IRestResponse response);

        private event WhenFailEventDelegate WhenFailDelegate;

        public GoogleEngine WhenFail(WhenFailEventDelegate action)
        {
            WhenFailDelegate += action;
            return this;
        }

        public delegate void WhenSucceedEventDelegate(GoogleSearchResults results, int startIdx);

        private event WhenSucceedEventDelegate WhenSucceedDelegate;

        public GoogleEngine WhenSucceed(WhenSucceedEventDelegate action)
        {
            WhenSucceedDelegate += action;
            return this;
        }

        public delegate void WhenCompletedEventDelegate();

        private event WhenCompletedEventDelegate WhenCompletedDelegate;

        public GoogleEngine WhenCompleted(WhenCompletedEventDelegate action)
        {
            WhenCompletedDelegate += action;
            return this;
        }

        // TODO: construct api by search options
        private string api =
            "https://www.googleapis.com/customsearch/v1?filter=1&cx=017201692871514580973%3Awwdg7q__" +
//            "mb4&imgSize=large&searchType=image&imgType=photo&safe=medium&key=AIzaSyCGcq3O8NN9U7YX-Pj3E7tZde0yaFFeUyY";
                "mb4&imgSize=large&searchType=image&imgType=photo&safe=medium&key=AIzaSyDQeqy9efF_ASgi2dk3Ortj2QNnz90RdOw";
//                "mb4&imgSize=large&searchType=image&imgType=photo&safe=medium&key=AIzaSyDXR8wBYL6al5jXIXTHpEF28CCuvL0fjKk";
//                "mb4&imgSize=large&searchType=image&imgType=photo&safe=medium&key=AIzaSyAur2Fc0ewRyGK0U8NCaaEfuY0g_sx-Qwk";
//                "mb4&imgSize=large&searchType=image&imgType=photo&safe=medium&key=AIzaSyArj45s-GLXKX8NSM6HGdSFtRvAMuKE2p0";

        public void Search(string query)
        {
            if (query.Trim().Length == 0) return;

            _lastTimeQuery = query;
            _nextStartIndex = NumOfItemsPerSearch;

            var barrier = CreateBarrier(NumOfItemsPerSearch/NumOfItemsPerRequest);

            for (var i = 0; i < NumOfItemsPerSearch; i += NumOfItemsPerRequest)
            {
                Search(query, i, barrier);
            }
        }

        public void SearchMore()
        {
            _nextStartIndex += NumOfItemsPerRequest;
            Search(_lastTimeQuery, _nextStartIndex, CreateBarrier(1));
        }

        private void Search(string query, int startIdx, Barrier barrier = null)
        {
            // TODO: construct api using options
            var restClient = new RestClient { BaseUrl = new Uri(api 
                + "&num=" + NumOfItemsPerRequest 
                + "&start=" + (startIdx + 1) 
                + "&q=" + query) };
            restClient.ExecuteAsync(new RestRequest(Method.GET), response =>
            {
                if (response.StatusCode != HttpStatusCode.OK
                    || response.Content.Trim().Length == 0)
                {
                    lock (_syncLock)
                    {
                        if (WhenFailDelegate != null && !_isFailedAlready)
                        {
                            // ensure only call failure delegate once
                            _isFailedAlready = true;
                            WhenFailDelegate(response);
                        }
                    }
                }
                else
                {
                    var deser = new JsonDeserializer();
                    var searchResults = deser.Deserialize<GoogleSearchResults>(response);
                    if (WhenSucceedDelegate != null)
                    {
                        WhenSucceedDelegate(searchResults, startIdx);
                    }
                }

                if (barrier != null)
                {
                    barrier.SignalAndWait();
                }
            });
        }

        private Barrier CreateBarrier(int numOfParticipants)
        {
            return new Barrier(numOfParticipants, b =>
            {
                _isFailedAlready = false;
                if (WhenCompletedDelegate != null)
                {
                    WhenCompletedDelegate();
                }
            });
        }
    }
}
