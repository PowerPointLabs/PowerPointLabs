using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.SearchEngine.VO;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;
using RestSharp;

namespace PowerPointLabs.ImageSearch.SearchEngine
{
    public partial class GoogleEngine
    {
        public const int NumOfItemsPerSearch = 30;
        public const int NumOfItemsPerRequest = 10;
        public const int MaxNumOfItems = 100;

        private const string GoogleCustomSearchApiBase = "https://www.googleapis.com";
        private const string GoogleCustomSearchApiResource = "/customsearch/v1";
        private const int SearchEngineTimeOut = 15000;

        private readonly IRestClient _restClient = new RestClient { Timeout = SearchEngineTimeOut };
        private readonly Dictionary<string, GoogleSearchResults> _params2ResultsMap = 
            new Dictionary<string, GoogleSearchResults>(); 

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
            _restClient.BaseUrl = GetBaseUrl();

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
            _restClient.BaseUrl = GetBaseUrl();

            Search(_lastTimeQuery, _nextStartIndex, CreateBarrier(1));
            _nextStartIndex += NumOfItemsPerRequest;
        }

        private void Search(string query, int startIdx, Barrier barrier)
        {
            var req = new RestRequest(Method.GET);
            AddBaseParameters(req);
            req.AddParameter("start", (startIdx + 1));
            req.AddParameter("q", query);

            var reqKey = GetKeyFromReq(req);
            // check if there's any cache
            if (_params2ResultsMap.ContainsKey(reqKey))
            {
                // barrier with multiple participants
                // require a task to run in
                new Task(() =>
                {
                    HandleSearchResults(barrier, () =>
                    {
                        var data = _params2ResultsMap[reqKey];
                        if (WhenSucceedDelegate != null)
                        {
                            WhenSucceedDelegate(data, startIdx);
                        }
                    });
                }).Start();
            }
            // go to search
            else
            {
                _restClient.ExecuteAsync<GoogleSearchResults>(req, response =>
                {
                    HandleSearchResults(barrier, () =>
                    {
                        TryHandleResponse(response, reqKey, startIdx);
                    });
                });
            }
        }
        # endregion

        # region Helper Funcs

        private delegate void TryHandleSearchResults();

        private void TryHandleResponse(IRestResponse<GoogleSearchResults> response, string reqKey, int startIdx)
        {
            if (response.StatusCode != HttpStatusCode.OK || StringUtil.IsEmpty(response.Content))
            {
                // ensure only call failure delegate once
                lock (_syncLock)
                {
                    if (WhenFailDelegate != null && !_isFailedAlready)
                    {
                        _isFailedAlready = true;
                        WhenFailDelegate(response);
                    }
                }
            }
            else // success
            {
                // add cache
                _params2ResultsMap.Add(reqKey, response.Data);
                if (WhenSucceedDelegate != null)
                {
                    WhenSucceedDelegate(response.Data, startIdx);
                }
            }
        }

        private void HandleSearchResults(Barrier barrier, TryHandleSearchResults tryHandle)
        {
            try
            {
                tryHandle();
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
        }

        /// <summary>
        /// manually add parameter for id and key in BaseUrl to avoid encoding
        /// </summary>
        /// <returns></returns>
        private Uri GetBaseUrl()
        {
            return new Uri(
                GoogleCustomSearchApiBase + 
                GoogleCustomSearchApiResource +
                "?cx=" + SearchOptions.SearchEngineId.Trim() +
                "&key=" + SearchOptions.ApiKey.Trim());
        }

        private void AddBaseParameters(IRestRequest req)
        {
            req.AddParameter("filter", "1");
            req.AddParameter("searchType", "image");
            req.AddParameter("safe", "medium");
            req.AddParameter("imgSize", SearchOptions.GetImageSize());
            req.AddParameter("imgType", SearchOptions.GetImageType());
            req.AddParameter("imgColorType", SearchOptions.GetColorType());
            if ("none" != SearchOptions.GetDominantColor())
            {
                req.AddParameter("imgDominantColor", SearchOptions.GetDominantColor());
            }
            if ("none" != SearchOptions.GetFileType())
            {
                req.AddParameter("fileType", SearchOptions.GetFileType());
            }
            req.AddParameter("num", NumOfItemsPerRequest);
        }

        private string GetKeyFromReq(IRestRequest req)
        {
            var strBuilder = new StringBuilder();
            var parameters = req.Parameters;
            foreach (var parameter in parameters)
            {
                strBuilder.Append(parameter);
                strBuilder.Append("&");
            }
            return strBuilder.ToString();
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
