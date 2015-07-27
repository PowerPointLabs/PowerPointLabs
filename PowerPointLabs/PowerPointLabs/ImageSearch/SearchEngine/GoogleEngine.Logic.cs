using System;
using System.Net;
using System.Threading;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.SearchEngine.VO;
using PowerPointLabs.ImageSearch.Util;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;
using RestSharp;
using RestSharp.Deserializers;

namespace PowerPointLabs.ImageSearch.SearchEngine
{
    public partial class GoogleEngine
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
        private bool _isFailedWithExceptionAlready;

        private SearchOptions SearchOptions { get; set; }

        # region APIs

        /// <exception cref="AssumptionFailedException">
        /// throw exception when options is null
        /// </exception>
        public GoogleEngine(SearchOptions options)
        {
            Assumption.Made(options != null, "options is null");
            SearchOptions = options;
        }

        /// <summary>
        /// Register the callbacks to handle search results
        /// </summary>
        /// <param name="query"></param>
        public void Search(string query)
        {
            if (StringUtil.IsEmpty(query)) return;

            _lastTimeQuery = query;
            _nextStartIndex = NumOfItemsPerSearch;

            var barrier = CreateBarrier(NumOfItemsPerSearch/NumOfItemsPerRequest);

            for (var i = 0; i < NumOfItemsPerSearch; i += NumOfItemsPerRequest)
            {
                Search(query, i, barrier);
            }
        }

        /// <summary>
        /// Register the callbacks to handle search results
        /// </summary>
        public void SearchMore()
        {
            if (StringUtil.IsEmpty(_lastTimeQuery)) return;

            Search(_lastTimeQuery, _nextStartIndex, CreateBarrier(1));
            _nextStartIndex += NumOfItemsPerRequest;
        }

        private void Search(string query, int startIdx, Barrier barrier = null)
        {
            var restClient = new RestClient
            {
                BaseUrl = new Uri(CreateApi() + "&num=" + NumOfItemsPerRequest 
                    + "&start=" + (startIdx + 1) + "&q=" + query) 
            };
            restClient.ExecuteAsync(new RestRequest(Method.GET), response =>
            {
                try
                {
                    if (response.StatusCode != HttpStatusCode.OK
                        || StringUtil.IsEmpty(response.Content))
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
                }
                catch (Exception e)
                {
                    lock (_syncLock)
                    {
                        if (WhenExceptionDelegate != null && !_isFailedWithExceptionAlready)
                        {
                            // ensure only call failure delegate once
                            _isFailedWithExceptionAlready = true;
                            try
                            {
                                WhenExceptionDelegate(e);
                            }
                            catch (Exception e2)
                            {
                                PowerPointLabsGlobals.Log("Error", e2.Message);
                            }
                        }
                    }
                }
                finally
                {
                    if (barrier != null)
                    {
                        barrier.SignalAndWait();
                    }
                }
            });
        }
        # endregion

        # region Helper Funcs
        private string CreateApi()
        {
            return "https://www.googleapis.com/customsearch/v1?filter=1&searchType=image&safe=medium"
                   + "&cx=" + SearchOptions.SearchEngineId.Trim()
                   + "&imgSize=" + SearchOptions.GetImageSize()
                   + "&imgType=" + SearchOptions.GetImageType()
                   + "&imgColorType=" + SearchOptions.GetColorType()
                   + ("none" != SearchOptions.GetDominantColor() ? "&imgDominantColor=" + SearchOptions.GetDominantColor() : "")
                   + ("none" != SearchOptions.GetFileType() ? "&fileType=" + SearchOptions.GetFileType() : "")
                   + "&key=" + SearchOptions.ApiKey.Trim();
        }

        private Barrier CreateBarrier(int numOfParticipants)
        {
            return new Barrier(numOfParticipants, b =>
            {
                var isSuccessful = !_isFailedAlready;
                _isFailedAlready = false;

                if (WhenCompletedDelegate == null) return;
                try
                {
                    WhenCompletedDelegate(isSuccessful);
                }
                catch (Exception e)
                {
                    PowerPointLabsGlobals.Log("Error", e.Message);
                }
            });
        }
        # endregion
    }
}
