using System;
using System.Diagnostics;
using PowerPointLabs.ActionFramework.Common.Logger;

namespace PowerPointLabs.ActionFramework.Common.Log
{
    public class Logger
    {
        public static void Log(string logText, LogType type = LogType.Info)
        {
            if (type.Equals(LogType.Info))
            {
                Trace.TraceInformation(DateTime.Now.ToString("yyyyMMddHHmmss") + ": " + logText);
            }
            else if (type.Equals(LogType.Error))
            {
                Trace.TraceError(DateTime.Now.ToString("yyyyMMddHHmmss") + ": " + logText);
            }
            else if (type.Equals(LogType.Warning))
            {
                Trace.TraceWarning(DateTime.Now.ToString("yyyyMMddHHmmss") + ": " + logText);
            }
        }
        public static void LogException(Exception e, string methodName)
        {
            Log(methodName + ": " + e.Message + " - " + e.GetType() + ": " + e.StackTrace, LogType.Error);
        }
    }
}
