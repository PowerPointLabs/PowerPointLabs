using System.Linq;
using System.Text.RegularExpressions;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.TagMatchers;

using Test.Util;

namespace Test.UnitTest.TagMatchers
{
    [TestClass]
    public class EndSpeedTests
    {
        private Regex tagRegex;

        [TestInitialize]
        public void Initialize()
        {
            tagRegex = new EndSpeedTagMatcher().Regex;
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchEndSpeed()
        {
            string testTag = "[EndSpeed]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchEndSpeedLowercase()
        {
            string testTag = "[endspeed]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchEndSpeedMixedCase()
        {
            string testTag = "[endSpeed]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchListSingleInSentence()
        {
            string sentence = "This is [speed: fast]a test[endspeed].";
            System.Collections.Generic.List<PowerPointLabs.Tags.ITag> matches = new EndSpeedTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 1, "More than one match.");

            PowerPointLabs.Tags.ITag match = matches[0];
            Assert.IsTrue(match.Start == 27, "Match start was incorrect.");
            Assert.IsTrue(match.End == 36, "Match end was incorrect.");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchListMultipleInSentence()
        {
            string sentence = "This [speed: slow]has multiple[endspeed][speed: fast] matches.[endspeed]";
            System.Collections.Generic.List<PowerPointLabs.Tags.ITag> matches = new EndSpeedTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 2, "Didn't match all.");
        }
    }
}
