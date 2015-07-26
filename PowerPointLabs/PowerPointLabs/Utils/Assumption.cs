using System;

namespace PowerPointLabs.Utils
{
    class Assumption
    {
        /// class T : Exception should have constructors:
        /// 1. that accepts assertMsg : string
        /// 2. that accepts assertMsg : string & e : Exception
        /// <exception>throw exception when condition is not matched</exception>
        public static void Made<T>(Func<bool> cond, string assertMsg) where T : Exception
        {
            try
            {
                if (!cond())
                {
                    throw (T) Activator.CreateInstance(typeof(T), assertMsg);
                }
            }
            catch (Exception e)
            {
                throw (T) Activator.CreateInstance(typeof(T), assertMsg, e);
            }
        }
    }
}
