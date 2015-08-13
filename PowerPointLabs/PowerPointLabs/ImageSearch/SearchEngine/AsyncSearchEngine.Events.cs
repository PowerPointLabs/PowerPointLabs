using System;
using RestSharp;

namespace PowerPointLabs.ImageSearch.SearchEngine
{
    public abstract partial class AsyncSearchEngine
    {
        public delegate void WhenExceptionEventDelegate(Exception e);

        public delegate void WhenFailEventDelegate(IRestResponse response);

        public delegate void WhenSucceedEventDelegate(object results, int startIdx);

        public delegate void WhenCompletedEventDelegate(bool isSuccessful);

        protected event WhenExceptionEventDelegate WhenExceptionDelegate;

        public AsyncSearchEngine WhenException(WhenExceptionEventDelegate action)
        {
            WhenExceptionDelegate += action;
            return this;
        }

        protected event WhenFailEventDelegate WhenFailDelegate;

        public AsyncSearchEngine WhenFail(WhenFailEventDelegate action)
        {
            WhenFailDelegate += action;
            return this;
        }

        protected event WhenSucceedEventDelegate WhenSucceedDelegate;

        public AsyncSearchEngine WhenSucceed(WhenSucceedEventDelegate action)
        {
            WhenSucceedDelegate += action;
            return this;
        }

        protected event WhenCompletedEventDelegate WhenCompletedDelegate;

        public AsyncSearchEngine WhenCompleted(WhenCompletedEventDelegate action)
        {
            WhenCompletedDelegate += action;
            return this;
        }
    }
}
