using System;

namespace PowerPointLabs.ResizeLab
{
    internal class ResizeLabUtil
    {
        /// <summary>
        /// Convert string to float. 
        /// Return null if the conversion is invalid.
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static float? ConvertToFloat(string input)
        {
            try
            {
                return float.Parse(input);
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// Check if the factor more than 0.
        /// </summary>
        /// <param name="factor"></param>
        /// <returns></returns>
        public static bool IsValidFactor(float? factor)
        {
            try
            {
                if (factor > 0)
                {
                    return true;
                }
            }
            catch (Exception)
            {
                return false;
            }
            return false;
        }
    }
}
