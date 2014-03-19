using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AudioGenTests
{
    public class TagTestHelpers
    {
        public static void MatchAndAssert(string testTag, Regex tagRegex)
        {
            var match = tagRegex.Match(testTag);

            Assert.IsTrue(match.Success, "Tag isn't matched.");
            Assert.IsTrue(match.Index == 0, "Match doesn't start at 0.");
            Assert.IsTrue(match.Length == testTag.Length, "Doesn't match entire tag.");
        }
    }
}