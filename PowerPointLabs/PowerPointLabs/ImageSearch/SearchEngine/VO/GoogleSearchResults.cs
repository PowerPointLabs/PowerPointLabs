using System.Collections.Generic;

namespace PowerPointLabs.ImageSearch.SearchEngine.VO
{
    public class GoogleSearchResults
    {
        public List<GoogleSearchResult> Items { get; set; }
    }

    public class GoogleSearchResult
    {
        public string Title { get; set; }
        public string Link { get; set; }
        public GoogleSearchResultImage Image { get; set; }
    }

    public class GoogleSearchResultImage
    {
        public string ContextLink { get; set; }
        public string Width { get; set; }
        public string Height { get; set; }
        public string ThumbnailLink { get; set; }
    }
}
