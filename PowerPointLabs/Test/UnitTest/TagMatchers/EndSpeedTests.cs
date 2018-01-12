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
            var testTag = "[EndSpeed]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchEndSpeedLowercase()
        {
            var testTag = "[endspeed]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchEndSpeedMixedCase()
        {
            var testTag = "[endSpeed]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchListSingleInSentence()
        {
            var sentence = "This is [speed: fast]a test[endspeed].";
            var matches = new EndSpeedTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 1, "More than one match.");

            var match = matches[0];
            Assert.IsTrue(match.Start == 27, "Match start was incorrect.");
            Assert.IsTrue(match.End == 36, "Match end was incorrect.");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchListMultipleInSentence()
        {
            var sentence = "This [speed: slow]has multiple[endspeed][speed: fast] matches.[endspeed]";
            var matches = new EndSpeedTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 2, "Didn't match all.");
        }
    }
}
