using System;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ActionFramework.Common.Logger;

namespace PowerPointLabs.Utils.Exceptions
{
    class AssumptionFailedException : Exception
    {
        public AssumptionFailedException(string errorMsg) : base(errorMsg)
        {
            Logger.Log(errorMsg, LogType.Error);
        }

        public AssumptionFailedException(string errorMsg, Exception e)
            : base(errorMsg, e)
        {
            Logger.Log(errorMsg, LogType.Error);
        }
    }
}
