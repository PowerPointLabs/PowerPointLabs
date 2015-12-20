using System;

namespace PowerPointLabs.Utils.Exceptions
{
    class AssumptionFailedException : Exception
    {
        public AssumptionFailedException(string errorMsg) : base(errorMsg)
        {
            PowerPointLabsGlobals.Log("Error", errorMsg);
        }

        public AssumptionFailedException(string errorMsg, Exception e)
            : base(errorMsg, e)
        {
            PowerPointLabsGlobals.Log("Error", errorMsg);
        }
    }
}
