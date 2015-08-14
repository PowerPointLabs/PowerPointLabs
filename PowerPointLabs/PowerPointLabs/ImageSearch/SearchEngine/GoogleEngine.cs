using System;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.ImageSearch.SearchEngine.VO;
using PowerPointLabs.Utils.Exceptions;
using RestSharp;
using RestSharp.Deserializers;

namespace PowerPointLabs.ImageSearch.SearchEngine
{
    public class GoogleEngine : AsyncSearchEngine
    {
        private const string GoogleCustomSearchApiBase = "https://www.googleapis.com";
        private const string GoogleCustomSearchApiResource = "/customsearch/v1";

        /// <exception cref="AssumptionFailedException">
        /// throw exception when options is null
        /// </exception>
        public GoogleEngine(SearchOptions options) : base(options) {}

        public static string Id() { return TextCollection.ImagesLabText.SearchEngineGoogle; }

        public override int NumOfItemsPerSearch() { return 30; }

        public override int NumOfItemsPerRequest() { return 10; }

        public override int MaxNumOfItems() { return 100; }

        /// <summary>
        /// manually add parameter for id and key in BaseUrl to avoid encoding
        /// </summary>
        /// <returns></returns>
        protected override Uri GetBaseUrl()
        {
            return new Uri(
                GoogleCustomSearchApiBase +
                GoogleCustomSearchApiResource +
                "?cx=" + SearchOptions.SearchEngineId.Trim() +
                "&key=" + SearchOptions.ApiKey.Trim());
        }

        protected override void AddBaseParameters(IRestRequest req)
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
            req.AddParameter("num", NumOfItemsPerRequest());
        }

        protected override void AddParametersPerSearch(RestRequest req, string query, int startIdx)
        {
            req.AddParameter("start", (startIdx + 1));
            req.AddParameter("q", query);
        }

        protected override object Deserialize(IRestResponse response)
        {
            var deser = new JsonDeserializer();
            var searchResults = deser.Deserialize<GoogleSearchResults>(response);
            return searchResults;
        }
    }
}
