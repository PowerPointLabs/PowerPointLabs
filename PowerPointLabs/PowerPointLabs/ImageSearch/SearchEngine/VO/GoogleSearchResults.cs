using System.Collections.Generic;

namespace PowerPointLabs.ImageSearch.SearchEngine.VO
{
    public class GoogleSearchResults
    {
        public List<SearchResult> Items { get; set; }
    }

    public class SearchResult
    {
        public string Title { get; set; }
        public string Link { get; set; }
        public SearchResultImage Image { get; set; }
    }

    public class SearchResultImage
    {
        public string ContextLink { get; set; }
        public string Width { get; set; }
        public string Height { get; set; }
        public string ThumbnailLink { get; set; }
    }
}
