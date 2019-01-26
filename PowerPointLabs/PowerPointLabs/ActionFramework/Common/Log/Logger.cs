using System;
using System.Diagnostics;

using PowerPointLabs.ActionFramework.Common.Logger;

namespace PowerPointLabs.ActionFramework.Common.Log
{
    public class Logger
    {
        private const string DateFormat = "yyyyMMddHHmmss";

        public static void Log(string logText, LogType type = LogType.Info)
        {
            if (type.Equals(LogType.Info))
            {
                Trace.TraceInformation(DateTime.Now.ToString(DateFormat) + ": " + logText);
            }
            else if (type.Equals(LogType.Error))
            {
                Trace.TraceError(DateTime.Now.ToString(DateFormat) + ": " + logText);
            }
            else if (type.Equals(LogType.Warning))
            {
                Trace.TraceWarning(DateTime.Now.ToString(DateFormat) + ": " + logText);
            }
        }
        public static void LogException(Exception e, string methodName)
        {
            if (e == null)
            {
                Log(methodName + ": " + "some error happened.", LogType.Error);
            }
            else
            {
                Log(methodName + ": " + e.Message + " - " + e.GetType() + ": " + e.StackTrace, LogType.Error);
            }
        }
    }
}
