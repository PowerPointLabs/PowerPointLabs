using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace PowerPointLabs.Utils
{
    public static class Common
    {
        public static int NextDefaultNumber(IEnumerable<string> nameList, Regex namePattern)
        {
            var defaultPattern = new Regex(@"^[^ ]+ (\d+)$");

            var temp = 0;
            var min = int.MaxValue;

            if (namePattern == null)
            {
                namePattern = defaultPattern;
            }

            foreach (var name in nameList)
            {
                if (namePattern.IsMatch(name))
                {
                    var currentCnt = int.Parse(namePattern.Match(name).Groups[1].Value);

                    if (currentCnt - temp != 1)
                    {
                        min = Math.Min(min, temp);
                    }

                    temp = currentCnt;
                }
            }

            if (min == int.MaxValue)
            {
                return temp + 1;
            }

            return min;
        }
    }
}