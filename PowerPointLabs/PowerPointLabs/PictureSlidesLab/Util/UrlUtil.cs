using System;
using System.Text.RegularExpressions;

using PowerPointLabs.PictureSlidesLab.Model;

using RestSharp.Contrib;

namespace PowerPointLabs.PictureSlidesLab.Util
{
    public class UrlUtil
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
                Uri googleImageUri = new Uri(url);
                System.Collections.Specialized.NameValueCollection parameters = HttpUtility.ParseQueryString(googleImageUri.Query);
                url = HttpUtility.UrlDecode(parameters.Get("imgurl"));
                item.ContextLink = HttpUtility.UrlDecode(parameters.Get("imgrefurl"));
                item.Source = item.ContextLink;
            }
        }
    }
}
