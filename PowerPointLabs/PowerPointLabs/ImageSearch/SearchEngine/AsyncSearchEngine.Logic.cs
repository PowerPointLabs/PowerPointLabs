using System;
using System.Collections.Generic;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Exceptions;
using RestSharp;

namespace PowerPointLabs.ImageSearch.SearchEngine
{
    public abstract partial class AsyncSearchEngine
    {
        // rest client
        private const int SearchEngineTimeOut = 15000; // 15 sec
        private readonly IRestClient _restClient = new RestClient { Timeout = SearchEngineTimeOut };

        // cache
        private readonly Dictionary<string, object> _params2ResultsMap = new Dictionary<string, object>();

        // states, used for Search More
        private string _lastTimeQuery = "";
        private int _nextStartIndex;

        // multi thread states
        private readonly Object _syncLock = new object();
        private bool _isFailedAlready;
        private bool _isFailedWithExceptionAlready;

        protected SearchOptions SearchOptions { get; set; }

        # region APIs

        /// <exception cref="AssumptionFailedException">
        /// throw exception when options is null
        /// </exception>
        protected AsyncSearchEngine(SearchOptions options)
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
            _nextStartIndex = NumOfItemsPerSearch();

            var barrier = CreateBarrier(NumOfItemsPerSearch()/NumOfItemsPerRequest());

            for (var i = 0; i < NumOfItemsPerSearch(); i += NumOfItemsPerRequest())
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
            _nextStartIndex += NumOfItemsPerRequest();
        }

        private void Search(string query, int startIdx, Barrier barrier)
        {
            var req = new RestRequest(Method.GET);
            AddBaseParameters(req);
            AddParametersPerSearch(req, query, startIdx);

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
                Authorize(_restClient);
                _restClient.BaseUrl = GetBaseUrl();
                _restClient.ExecuteAsync(req, response =>
                {
                    HandleSearchResults(barrier, () =>
                    {
                        TryHandleResponse(response, reqKey, startIdx);
                    });
                });
            }
        }

        # endregion

        # region To be overridden

        protected virtual void Authorize(IRestClient client) { }

        public abstract int NumOfItemsPerSearch();

        public abstract int NumOfItemsPerRequest();

        public abstract int MaxNumOfItems();

        protected abstract Uri GetBaseUrl();

        protected abstract void AddBaseParameters(IRestRequest req);

        protected abstract void AddParametersPerSearch(RestRequest req, string query, int startIdx);

        protected abstract object Deserialize(IRestResponse response);

        # endregion

        # region Helper Funcs

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

        private delegate void TryHandleSearchResults();

        private void TryHandleResponse(IRestResponse response, string reqKey, int startIdx)
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
                var data = Deserialize(response);
                // add cache
                _params2ResultsMap.Add(reqKey, data);
                if (WhenSucceedDelegate != null)
                {
                    WhenSucceedDelegate(data, startIdx);
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
