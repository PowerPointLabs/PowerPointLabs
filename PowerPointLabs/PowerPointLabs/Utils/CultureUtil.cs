using System;
using System.Globalization;
using System.Reflection;
using System.Threading;

namespace PowerPointLabs.Utils
{
    public class CultureUtil
    {
        private static CultureInfo _originalCultureInfo = Thread.CurrentThread.CurrentCulture;

        /// <summary>
        /// Taken from http://blog.rastating.com/setting-default-currentculture-in-all-versions-of-net/
        /// in order to fix culture settings issue, e.g. 1,1 in Italy and 1.1 in US. 
        /// </summary>
        /// <param name="culture"></param>
        public static void SetDefaultCulture(CultureInfo culture)
        {
            Type type = typeof(CultureInfo);

            try
            {
                type.InvokeMember("s_userDefaultCulture",
                    BindingFlags.SetField | BindingFlags.NonPublic | BindingFlags.Static,
                    null,
                    culture,
                    new object[] {culture});

                type.InvokeMember("s_userDefaultUICulture",
                    BindingFlags.SetField | BindingFlags.NonPublic | BindingFlags.Static,
                    null,
                    culture,
                    new object[] {culture});
            }
            catch
            {
                // this version of .NET doesn't have this field
            }

            try
            {
                type.InvokeMember("m_userDefaultCulture",
                    BindingFlags.SetField | BindingFlags.NonPublic | BindingFlags.Static,
                    null,
                    culture,
                    new object[] {culture});

                type.InvokeMember("m_userDefaultUICulture",
                    BindingFlags.SetField | BindingFlags.NonPublic | BindingFlags.Static,
                    null,
                    culture,
                    new object[] {culture});
            }
            catch
            {
                // this version of .NET doesn't have this field
            }
        }

        public static CultureInfo GetOriginalCulture()
        {
            return _originalCultureInfo;
        }
    }
}
