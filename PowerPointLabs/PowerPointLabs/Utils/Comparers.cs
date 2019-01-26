using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace PowerPointLabs.Utils
{
    public static class Comparers
    {
        public class AtomicNumberStringCompare : IComparer<string>
        {
            public int Compare(string thisString, string otherString)
            {
                // some characters + number
                Regex pattern = new Regex(@"([^\d]+)(\d+)");
                Match thisStringMatch = pattern.Match(thisString);
                Match otherStringMatch = pattern.Match(otherString);

                // specially compare the pattern, after run out of the pattern, compare
                // 2 strings normally
                while (thisStringMatch.Success &&
                       otherStringMatch.Success)
                {
                    string thisStringPart = thisStringMatch.Groups[1].Value;
                    int thisNumPart = int.Parse(thisStringMatch.Groups[2].Value);

                    string otherStringPart = otherStringMatch.Groups[1].Value;
                    int otherNumPart = int.Parse(otherStringMatch.Groups[2].Value);

                    // if string part is not the same, we can tell the diff
                    if (!string.Equals(thisStringPart, otherStringPart))
                    {
                        break;
                    }

                    // if string part is the same but number part is different, we can
                    // tell the diff
                    if (thisNumPart != otherNumPart)
                    {
                        return thisNumPart - otherNumPart;
                    }

                    // two parts are identical, find next match
                    thisStringMatch = thisStringMatch.NextMatch();
                    otherStringMatch = otherStringMatch.NextMatch();
                }

                // case sensitive comparing, invariant for cultures
                return string.Compare(thisString, otherString, false, CultureInfo.InvariantCulture);
            }
        }
    }
}
