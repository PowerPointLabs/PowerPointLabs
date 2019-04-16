using System.Linq;
using System.Text.RegularExpressions;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.TagMatchers;

using Test.Util;

namespace Test.UnitTest.TagMatchers
{
    [TestClass]
    public class PauseTest
    {
        private Regex tagRegex;

        [TestInitialize]
        public void Initialize()
        {
            tagRegex = new PauseTagMatcher().Regex;
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchIntegerInterval()
        {
            string testTag = "[Pause: 2]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchMultipleDigitIntegerInterval()
        {
            string testTag = "[Pause: 23]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchIntegerIntervalLowercase()
        {
            string testTag = "[pause: 2]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchMultipleDigitIntegerIntervalLowercase()
        {
            string testTag = "[pause: 23]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchDecimalInterval()
        {
            string testTag = "[Pause: 2.5]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void DontMatchMultipleDecimals()
        {
            string testTag = "[Pause: 2.5.1]";
            Assert.IsFalse(tagRegex.IsMatch(testTag), "Matched multiple decimals.");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchSingleInSentence()
        {
            string sentence = "This has a pause [pause: 2] right here.";
            System.Collections.Generic.List<PowerPointLabs.Tags.ITag> matches = new PauseTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 1, "More than one match.");

            PowerPointLabs.Tags.ITag match = matches[0];
            Assert.IsTrue(match.Start == 17, "Match start was incorrect.");
            Assert.IsTrue(match.End == 26, "Match end was incorrect.");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchMultipleInSentence()
        {
            string sentence = "This has [pause: 2] many [pause: 2.4] pauses.";
            System.Collections.Generic.List<PowerPointLabs.Tags.ITag> matches = new PauseTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 2, "Didn't match all.");
        }
    }
}
