using System;

namespace PowerPointLabs.ImageSearch.Domain.Exceptions
{
    class ImagesLabException : Exception
    {
        public ImagesLabException(string errorMsg) : base(errorMsg)
        {
            PowerPointLabsGlobals.Log("Error", errorMsg);
        }

        public ImagesLabException(string errorMsg, Exception e)
            : base(errorMsg, e)
        {
            PowerPointLabsGlobals.Log("Error", errorMsg);
        }
    }
}
