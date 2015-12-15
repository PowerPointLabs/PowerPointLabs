using PowerPointLabs.ImageSearch.SearchEngine.VO;

namespace PowerPointLabs.ImageSearch.Util
{
    class VOUtil
    {
        public static int GetCount(object results)
        {
            // assume not null
            if (results is BingSearchResults)
            {
                return (results as BingSearchResults).D.Results.Count;
            }
            if (results is GoogleSearchResults)
            {
                return (results as GoogleSearchResults).Items.Count;
            }
            return 0;
        }

        public static object GetItem(object results, int index)
        {
            // assume not null
            if (results is BingSearchResults && index < GetCount(results))
            {
                return (results as BingSearchResults).D.Results[index];
            }
            if (results is GoogleSearchResults && index < GetCount(results))
            {
                return (results as GoogleSearchResults).Items[index];
            }
            return null;
        }

        public static string GetTitle(object result)
        {
            if (result is BingSearchResult)
            {
                return (result as BingSearchResult).Title;
            }
            if (result is GoogleSearchResult)
            {
                return (result as GoogleSearchResult).Title;
            }
            return "";
        }

        public static string GetLink(object result)
        {
            if (result is BingSearchResult)
            {
                return (result as BingSearchResult).MediaUrl;
            }
            if (result is GoogleSearchResult)
            {
                return (result as GoogleSearchResult).Link;
            }
            return "";
        }

        public static string GetContextLink(object result)
        {
            if (result is BingSearchResult)
            {
                return (result as BingSearchResult).SourceUrl;
            }
            if (result is GoogleSearchResult)
            {
                return (result as GoogleSearchResult).Image.ContextLink;
            }
            return "";
        }

        public static string GetHeight(object result)
        {
            if (result is BingSearchResult)
            {
                return (result as BingSearchResult).Height;
            }
            if (result is GoogleSearchResult)
            {
                return (result as GoogleSearchResult).Image.Height;
            }
            return "";
        }

        public static string GetWidth(object result)
        {
            if (result is BingSearchResult)
            {
                return (result as BingSearchResult).Width;
            }
            if (result is GoogleSearchResult)
            {
                return (result as GoogleSearchResult).Image.Width;
            }
            return "";
        }

        public static string GetThumbnailLink(object result)
        {
            if (result is BingSearchResult)
            {
                return (result as BingSearchResult).Thumbnail.MediaUrl;
            }
            if (result is GoogleSearchResult)
            {
                return (result as GoogleSearchResult).Image.ThumbnailLink;
            }
            return "";
        }
    }
}
