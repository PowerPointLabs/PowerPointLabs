using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;

namespace PowerPointLabs.Utils
{
    public static class CommonUtil
    {
        #region text
        /// <summary>
        /// Used for the method UniqueDigitString().
        /// _sessionGlobalUniqueIndex is guaranteed to be unique within the same powerpoint session.
        /// </summary>
        private static int _sessionGlobalUniqueIndex = 0;

        public static string NextAvailableName(List<string> nameList, string name)
        {
            List<string> orderedNameList = nameList.OrderBy(item => item, new Comparers.AtomicNumberStringCompare()).ToList();
            int nextDefaultNumber = NextDefaultNumber(orderedNameList, name);

            return nextDefaultNumber == 0 ? name : string.Format("{0} {1}", name, nextDefaultNumber);
        }

        public static string SkipRegexCharacter(string str)
        {
            Regex replacePattern = new Regex(@"([^\d\s\w])");

            return replacePattern.Replace(str, "\\$1");
        }

        /// <summary>
        /// Used to encode a string to make it a safe file name.
        /// Base64 usually uses a / and a + character. This uses a _ and a - instead. (safe for file names)
        /// </summary>
        public static string FilenameBase64(string str)
        {
            byte[] plainTextBytes = System.Text.Encoding.UTF8.GetBytes(str);
            str = Convert.ToBase64String(plainTextBytes);
            str = str.Replace("+", "-");
            str = str.Replace("/", "_");
            return str;
        }

        /// <summary>
        /// Used to encode a string into base64. Uses only alphanumeric chars and + and / characters.
        /// </summary>
        public static string Base64Encode(string str)
        {
            byte[] plainTextBytes = System.Text.Encoding.UTF8.GetBytes(str);
            str = Convert.ToBase64String(plainTextBytes);
            return str;
        }

        /// <summary>
        /// Used to decode a string from base64. Uses only alphanumeric chars and + and / characters.
        /// </summary>
        public static string Base64Decode(string base64EncodedData)
        {
            byte[] base64EncodedBytes = Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }

        /// <summary>
        /// Converts a list of strings into a single string for storage.
        /// Adds a checksum at the end as well
        /// </summary>
        public static string SerializeCollection(List<string> collection)
        {
            string serialized = string.Join("@", collection.Select(Base64Encode));
            return serialized + "@" + ComputeCheckSum(serialized);
        }

        /// <summary>
        /// Converts a single serialised string back into a list of strings. (strings serialized by SerializeCollection)
        /// Verifies the checksum as well. If checksum does not match, returns null.
        /// </summary>
        public static List<string> UnserializeCollection(string dataString)
        {
            int lastDelim = dataString.LastIndexOf('@');
            if (lastDelim == -1)
            {
                return null;
            }

            // Verify checksum
            string hashCode = dataString.Substring(lastDelim + 1);
            string serialized = dataString.Substring(0, lastDelim);
            if (ComputeCheckSum(serialized).ToString() != hashCode)
            {
                return null;
            }

            return serialized.Split(new[] {'@'}, StringSplitOptions.None)
                             .Select(Base64Decode)
                             .ToList();
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

        /// <summary>
        /// Returns a list of (nStrings) strings that differ from all strings in presentStrings
        /// </summary>
        public static string[] GetUnusedStrings(IEnumerable<string> presentStrings, int nStrings)
        {
            List<string> unusedStrings = new List<string>();

            int longestString = presentStrings.Select(str => str.Length).Max();
            string baseString = new string('a', longestString);
            
            for (int i = 0; i < nStrings; ++i)
            {
                unusedStrings.Add(baseString + i);
            }
            return unusedStrings.ToArray();
        }

        public static string SplitCamelCase(string input)
        {
            return Regex.Replace(input, "([A-Z])", " $1", RegexOptions.Compiled).Trim();
        }
        #endregion // text

        #region Math

        /// <summary>
        /// Computes ceil(dividend / divisor)
        /// </summary>
        public static int CeilingDivide(int dividend, int divisor)
        {
            return (dividend + divisor - 1) / divisor;
        }

        public static PointF RotatePoint(PointF p, PointF origin, float rotation)
        {
            double rotationInRadian = DegreeToRadian(rotation);
            double rotatedX = Math.Cos(rotationInRadian) * (p.X - origin.X) - Math.Sin(rotationInRadian) * (p.Y - origin.Y) + origin.X;
            double rotatedY = Math.Sin(rotationInRadian) * (p.X - origin.X) + Math.Cos(rotationInRadian) * (p.Y - origin.Y) + origin.Y;

            return new PointF((float)rotatedX, (float)rotatedY);
        }

        public static double DegreeToRadian(float angle)
        {
            return angle / 180.0 * Math.PI;
        }

        public static float Wrap(float value, float lower, float upper)
        {
            float range = upper - lower;
            float zeroBasedValue = value - lower;
            float result = ((zeroBasedValue + range) % range) + lower;
            return result;
        }

        #endregion

        #region Helper Function
        private static int NextDefaultNumber(List<string> nameList, string name)
        {
            Regex namePattern = new Regex(string.Format("{0}(?: (\\d+))*", name));
            int min = 0;

            if (!string.IsNullOrEmpty(namePattern.Match(nameList[0]).Groups[1].Value))
            {
                return min;
            }

            for (int i = 1; i < nameList.Count; i++)
            {
                int currentCnt = int.Parse(namePattern.Match(nameList[i]).Groups[1].Value);

                if (currentCnt != min + 1)
                {
                    return min + 1;
                }

                min++;
            }

            return min + 1;
        }

        /// <summary>
        /// Works like a hashcode function, but returns a digit string.
        /// Except that hashcode isn't consistent across all platforms / implementations. This is.
        /// </summary>
        private static string ComputeCheckSum(string s)
        {
            uint x = 0;
            foreach (char c in s)
            {
                x += c;
                x *= 565325351;
            }
            return x.ToString();
        }
        #endregion
    }
}
