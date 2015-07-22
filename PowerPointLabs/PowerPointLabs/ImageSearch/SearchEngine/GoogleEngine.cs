using System;
using System.Net;
using System.Threading;
using PowerPointLabs.ImageSearch.Model;
using PowerPointLabs.ImageSearch.SearchEngine.VO;
using RestSharp;
using RestSharp.Deserializers;

namespace PowerPointLabs.ImageSearch.SearchEngine
{
    public class GoogleEngine
    {
        public const int NumOfItemsPerSearch = 30;
        public const int NumOfItemsPerRequest = 10;
        public const int MaxNumOfItems = 100;

        // state, used for Search More
        private string _lastTimeQuery = "";
        private int _nextStartIndex;

        // multi thread state
        private readonly Object _syncLock = new object();
        private bool _isFailedAlready;

        public SearchOptions SearchOptions { get; private set; }

        public GoogleEngine(SearchOptions options)
        {
            SearchOptions = options;
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

        public delegate void WhenCompletedEventDelegate(bool isSuccessful);

        private event WhenCompletedEventDelegate WhenCompletedDelegate;

        public GoogleEngine WhenCompleted(WhenCompletedEventDelegate action)
        {
            WhenCompletedDelegate += action;
            return this;
        }

        private string GetApi()
        {
            return "https://www.googleapis.com/customsearch/v1?filter=1&searchType=image&safe=medium"
                   + "&cx=" + SearchOptions.SearchEngineId.Trim()
                   + "&imgSize=" + SearchOptions.GetImageSize()
                   + "&imgType=" + SearchOptions.GetImageType()
                   + "&imgColorType=" + SearchOptions.GetColorType()
                   + ("none" != SearchOptions.GetDominantColor()? "&imgDominantColor=" + SearchOptions.GetDominantColor() : "")
                   + "&key=" + SearchOptions.ApiKey.Trim();
        }

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
            Search(_lastTimeQuery, _nextStartIndex, CreateBarrier(1));
            _nextStartIndex += NumOfItemsPerRequest;
        }

        private void Search(string query, int startIdx, Barrier barrier = null)
        {
            var restClient = new RestClient
            {
                BaseUrl = new Uri(GetApi() 
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
                var isSuccessful = !_isFailedAlready;
                _isFailedAlready = false;
                if (WhenCompletedDelegate != null)
                {
                    WhenCompletedDelegate(isSuccessful);
                }
            });
        }
    }
}
