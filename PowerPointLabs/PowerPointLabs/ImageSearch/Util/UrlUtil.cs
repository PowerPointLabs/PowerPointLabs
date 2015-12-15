using System;
using System.Text.RegularExpressions;
using PowerPointLabs.ImageSearch.Domain;
using RestSharp.Contrib;

namespace PowerPointLabs.ImageSearch.Util
{
    class UrlUtil
    {
        /// <summary>
        /// taken from http://stackoverflow.com/a/5717342
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static bool IsUrlValid(string url)
        {
            return Uri.IsWellFormedUriString(url, UriKind.Absolute);
        }

        public static bool IsValidGoogleImageLink(string url)
        {
            return new Regex(@".*google.*imgres.*imgurl=.*imgrefurl=.*", 
                RegexOptions.Compiled | RegexOptions.IgnoreCase)
                .IsMatch(url);
        }

        public static void GetMetaInfo(ref string url, ImageItem item)
        {
            if (IsValidGoogleImageLink(url))
            {
                var googleImageUri = new Uri(url);
                var parameters = HttpUtility.ParseQueryString(googleImageUri.Query);
                url = parameters.Get("imgurl");
                item.FullSizeImageUri = url;
                item.ContextLink = parameters.Get("imgrefurl");
            }
        }
    }
}
