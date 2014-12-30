using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PowerPointLabs.Utils
{
    public static class Common
    {
        public static string NextAvailableName(List<string> nameList, string name)
        {
            var orderedNameList = nameList.OrderBy(item => item, new Comparers.AtomicNumberStringCompare()).ToList();
            var nextDefaultNumber = NextDefaultNumber(orderedNameList, name);

            return string.Format("{0} {1}", name, nextDefaultNumber);
        }

        public static string SkipRegexCharacter(string str)
        {
            var replacePattern = new Regex(@"([^\d\s\w])");

            return replacePattern.Replace(str, "\\$1");
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