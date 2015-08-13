using System;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.SearchEngine.VO;
using PowerPointLabs.Utils.Exceptions;
using RestSharp;
using RestSharp.Deserializers;

namespace PowerPointLabs.ImageSearch.SearchEngine
{
    class BingEngine : AsyncSearchEngine
    {
        private const string BingApiBase = "https://api.datamarket.azure.com";
        private const string BingApiResource = "/Bing/Search/v1/Image";

        /// <exception cref="AssumptionFailedException">
        /// throw exception when options is null
        /// </exception>
        public BingEngine(SearchOptions options) : base(options) {}

        public static string Id() { return TextCollection.ImagesLabText.SearchEngineBing; }

        public override int NumOfItemsPerSearch() { return 50; }

        public override int NumOfItemsPerRequest() { return 50; }

        public override int MaxNumOfItems() { return 1000; }

        protected override void Authorize(IRestClient client)
        {
            client.Authenticator = new HttpBasicAuthenticator("", SearchOptions.BingApiKey.Trim());
        }

        protected override Uri GetBaseUrl()
        {
            return new Uri(BingApiBase + BingApiResource);
        }

        protected override void AddBaseParameters(IRestRequest req)
        {
            req.AddParameter("Adult", AddQuotes("Moderate"));
            req.AddParameter("Options", AddQuotes("DisableLocationDetection"));

            req.AddParameter("ImageFilters", AddQuotes(SearchOptions.GetBingImageFilters()));

            req.AddParameter("$format", "json");
            req.AddParameter("$top", NumOfItemsPerRequest());
        }

        protected override void AddParametersPerSearch(RestRequest req, string query, int startIdx)
        {
            req.AddParameter("$skip", startIdx);
            req.AddParameter("Query", AddQuotes(query));
        }

        protected override object Deserialize(IRestResponse response)
        {
            var deser = new JsonDeserializer();
            var searchResults = deser.Deserialize<BingSearchResults>(response);
            return searchResults;
        }

        private string AddQuotes(string param)
        {
            return "'" + param + "'";
        }
    }
}
