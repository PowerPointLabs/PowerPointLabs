using System.Text.RegularExpressions;

using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Test.Util
{
    public class TagUtil
    {
        public static void MatchAndAssert(string testTag, Regex tagRegex)
        {
            Match match = tagRegex.Match(testTag);

            Assert.IsTrue(match.Success, "Tag isn't matched.");
            Assert.IsTrue(match.Index == 0, "Match doesn't start at 0.");
            Assert.IsTrue(match.Length == testTag.Length, "Doesn't match entire tag.");
        }
    }
}