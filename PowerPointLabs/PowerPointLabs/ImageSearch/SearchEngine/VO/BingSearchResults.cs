using System.Collections.Generic;

namespace PowerPointLabs.ImageSearch.SearchEngine.VO
{
    public class BingSearchResults
    {
        public BingSearchResultsWrapper D { get; set; }
    }

    public class BingSearchResultsWrapper
    {
        public List<BingSearchResult> Results { get; set; }
    }

    public class BingSearchResult
    {
        public string Title { get; set; }
        public string MediaUrl { get; set; }
        public string SourceUrl { get; set; }
        public string Width { get; set; }
        public string Height { get; set; }
        public BingThumbnail Thumbnail { get; set; }
    }

    public class BingThumbnail
    {
        public string MediaUrl { get; set; }
    }
}
