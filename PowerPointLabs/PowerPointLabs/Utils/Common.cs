using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace PowerPointLabs.Utils
{
    public static class Common
    {
        public static string NextAvailableName(List<string> nameList, Regex namePattern)
        {
            var nameFormat = string.Empty;
            var substitueString = string.Format("$1 {0}", NextDefaultNumber(nameList, namePattern, ref nameFormat));

            return namePattern.Replace(nameFormat, substitueString);
        }

        public static string SkipRegexCharacter(string str)
        {
            var replacePattern = new Regex(@"([^\d\s\w])");

            return replacePattern.Replace(str, "\\$1");
        }

        # region Helper Function
        private static int NextDefaultNumber(IEnumerable<string> nameList, Regex namePattern, ref string nameFormat)
        {
            var defaultPattern = new Regex(@"^(\.+) (\d+)$");

            var temp = 0;
            var min = int.MaxValue;

            if (namePattern == null)
            {
                namePattern = defaultPattern;
            }

            foreach (var name in nameList.Where(name => namePattern.IsMatch(name)))
            {
                if (nameFormat == string.Empty)
                {
                    nameFormat = namePattern.Match(name).Groups[1].Value;
                }

                var currentCnt = int.Parse(namePattern.Match(name).Groups[2].Value);

                if (currentCnt - temp != 1)
                {
                    min = Math.Min(min, temp);
                }

                temp = currentCnt;
            }

            if (min == int.MaxValue)
            {
                return temp + 1;
            }

            return min;
        }
        # endregion
    }
}