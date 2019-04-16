using System;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ActionFramework.Common.Logger;

namespace PowerPointLabs.CropLab
{
    class CropLabOperationException : Exception
    {
        public CropLabOperationException(string errorMsg) : base(errorMsg)
        {
            Logger.Log(errorMsg, LogType.Error);
        }

        public CropLabOperationException(string errorMsg, Exception e)
            : base(errorMsg, e)
        {
            Logger.Log(errorMsg, LogType.Error);
        }
    }
}
