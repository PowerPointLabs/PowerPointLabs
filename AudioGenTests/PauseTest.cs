using System.Linq;
using System.Text.RegularExpressions;
using AudioGen.TagMatchers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AudioGenTests
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
        public void MatchIntegerInterval()
        {
            var testTag = "[Pause: 2]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchMultipleDigitIntegerInterval()
        {
            var testTag = "[Pause: 23]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchIntegerIntervalLowercase()
        {
            var testTag = "[pause: 2]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchMultipleDigitIntegerIntervalLowercase()
        {
            var testTag = "[pause: 23]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchDecimalInterval()
        {
            var testTag = "[Pause: 2.5]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void DontMatchMultipleDecimals()
        {
            var testTag = "[Pause: 2.5.1]";
            Assert.IsFalse(tagRegex.IsMatch(testTag), "Matched multiple decimals.");
        }

        [TestMethod]
        public void MatchSingleInSentence()
        {
            var sentence = "This has a pause [pause: 2] right here.";
            var matches = new PauseTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 1, "More than one match.");

            var match = matches[0];
            Assert.IsTrue(match.Start == 17, "Match start was incorrect.");
            Assert.IsTrue(match.End == 26, "Match end was incorrect.");
        }

        [TestMethod]
        public void MatchMultipleInSentence()
        {
            var sentence = "This has [pause: 2] many [pause: 2.4] pauses.";
            var matches = new PauseTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 2, "Didn't match all.");
        }
    }
}
