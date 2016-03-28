using System;

namespace PowerPointLabs.ResizeLab
{
    internal class ResizeLabUtil
    {
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
