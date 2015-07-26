using System;
using PowerPointLabs.ImageSearch.SearchEngine.VO;
using RestSharp;

namespace PowerPointLabs.ImageSearch.SearchEngine
{
    partial class GoogleEngine
    {
        public delegate void WhenExceptionEventDelegate(Exception e);

        private event WhenExceptionEventDelegate WhenExceptionDelegate;

        public GoogleEngine WhenException(WhenExceptionEventDelegate action)
        {
            WhenExceptionDelegate += action;
            return this;
        }

        public delegate void WhenFailEventDelegate(IRestResponse response);

        private event WhenFailEventDelegate WhenFailDelegate;

        public GoogleEngine WhenFail(WhenFailEventDelegate action)
        {
            WhenFailDelegate += action;
            return this;
        }

        public delegate void WhenSucceedEventDelegate(GoogleSearchResults results, int startIdx);

        private event WhenSucceedEventDelegate WhenSucceedDelegate;

        public GoogleEngine WhenSucceed(WhenSucceedEventDelegate action)
        {
            WhenSucceedDelegate += action;
            return this;
        }

        public delegate void WhenCompletedEventDelegate(bool isSuccessful);

        private event WhenCompletedEventDelegate WhenCompletedDelegate;

        public GoogleEngine WhenCompleted(WhenCompletedEventDelegate action)
        {
            WhenCompletedDelegate += action;
            return this;
        }
    }
}
