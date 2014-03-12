using System.Linq;
using System.Text.RegularExpressions;
using AudioGen.TagMatchers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AudioGenTests
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
        public void MatchEndSpeed()
        {
            var testTag = "[EndSpeed]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchEndSpeedLowercase()
        {
            var testTag = "[endspeed]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchEndSpeedMixedCase()
        {
            var testTag = "[endSpeed]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
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
        public void MatchListMultipleInSentence()
        {
            var sentence = "This [speed: slow]has multiple[endspeed][speed: fast] matches.[endspeed]";
            var matches = new EndSpeedTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 2, "Didn't match all.");
        }
    }
}
