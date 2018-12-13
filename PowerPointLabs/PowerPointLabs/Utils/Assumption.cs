using System;

using PowerPointLabs.Utils.Exceptions;

namespace PowerPointLabs.Utils
{
    class Assumption
    {
        /// <exception cref="AssumptionFailedException">
        /// throw exception when condition is not matched
        /// </exception>
        public static void Made(bool cond, string assertMsg)
        {
            try
            {
                if (!cond)
                {
                    throw new AssumptionFailedException(assertMsg);
                }
            }
            catch (Exception e)
            {
                throw new AssumptionFailedException(assertMsg, e);
            }
        }
    }
}
