using System.Collections.Generic;

namespace PowerPointLabs.ImageSearch.VO
{
    class SearchResults
    {
        public List<SearchResult> Items { get; set; }
    }

    class SearchResult
    {
        public string Title { get; set; }
        public string Link { get; set; }
        public SearchResultImage Image { get; set; }
    }

    class SearchResultImage
    {
        public string ContextLink { get; set; }
        public string Width { get; set; }
        public string Height { get; set; }
        public string ThumbnailLink { get; set; }
    }
}
