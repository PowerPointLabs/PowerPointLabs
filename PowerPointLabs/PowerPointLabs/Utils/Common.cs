using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PowerPointLabs.Utils
{
    public static class Common
    {
        /// <summary>
        /// Used for the method UniqueDigitString().
        /// _sessionGlobalUniqueIndex is guaranteed to be unique within the same powerpoint session.
        /// </summary>
        private static int _sessionGlobalUniqueIndex = 0;

        public static string NextAvailableName(List<string> nameList, string name)
        {
            var orderedNameList = nameList.OrderBy(item => item, new Comparers.AtomicNumberStringCompare()).ToList();
            var nextDefaultNumber = NextDefaultNumber(orderedNameList, name);

            return nextDefaultNumber == 0 ? name : string.Format("{0} {1}", name, nextDefaultNumber);
        }

        public static string SkipRegexCharacter(string str)
        {
            var replacePattern = new Regex(@"([^\d\s\w])");

            return replacePattern.Replace(str, "\\$1");
        }

        /// <summary>
        /// Used to encode a string to make it a safe file name.
        /// Base64 usually uses a / and a + character. This uses a _ and a - instead. (safe for file names)
        /// </summary>
        public static string FilenameBase64(string str)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(str);
            str = Convert.ToBase64String(plainTextBytes);
            str = str.Replace("+", "-");
            str = str.Replace("/", "_");
            return str;
        }

        /// <summary>
        /// Generates a unique string of digits to be used in slide names.
        /// _sessionGlobalUniqueIndex is guaranteed to be unique within the same powerpoint session.
        /// DateTimeNow.ToString guaranteed to be unique across different powerpoint sessions.
        /// </summary>
        public static string UniqueDigitString()
        {
            string digitString = DateTime.Now.ToString("yyyyMMddHHmmssffff") + _sessionGlobalUniqueIndex;
            _sessionGlobalUniqueIndex++;
            return digitString;
        }

        # region Helper Function
        private static int NextDefaultNumber(List<string> nameList, string name)
        {
            var namePattern = new Regex(string.Format("{0}(?: (\\d+))*", name));
            var min = 0;

            if (!string.IsNullOrEmpty(namePattern.Match(nameList[0]).Groups[1].Value))
            {
                return min;
            }

            for (var i = 1; i < nameList.Count; i ++)
            {
                var currentCnt = int.Parse(namePattern.Match(nameList[i]).Groups[1].Value);

                if (currentCnt != min + 1)
                {
                    return min + 1;
                }

                min++;
            }

            return min + 1;
        }
        # endregion
    }
}