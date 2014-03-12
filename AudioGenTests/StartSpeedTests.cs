using System.Linq;
using System.Text.RegularExpressions;
using AudioGen.TagMatchers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AudioGenTests
{
    [TestClass]
    public class StartSpeedTests
    {
        private Regex tagRegex;

        [TestInitialize]
        public void Initialize()
        {
            tagRegex = new StartSpeedTagMatcher().Regex;
        }

        [TestMethod]
        public void MatchExtraSlowLowerCase()
        {
            var testTag = "[speed: extra slow]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchExtraSlowUpperCase()
        {
            var testTag = "[Speed: Extra Slow]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchExtraSlowMixedCase()
        {
            var testTag = "[Speed: extra Slow]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchSlowLowerCase()
        {
            var testTag = "[speed: slow]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchSlowUpperCase()
        {
            var testTag = "[Speed: Slow]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchSlowMixedCase()
        {
            var testTag = "[Speed: slow]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchMediumLowerCase()
        {
            var testTag = "[speed: medium]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchMediumUpperCase()
        {
            var testTag = "[Speed: Medium]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchMediumMixedCase()
        {
            var testTag = "[Speed: medium]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchFastLowerCase()
        {
            var testTag = "[speed: fast]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchFastUpperCase()
        {
            var testTag = "[Speed: Fast]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchFastMixedCase()
        {
            var testTag = "[Speed: fast]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchExtraFastLowerCase()
        {
            var testTag = "[speed: extra fast]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchExtraFastUpperCase()
        {
            var testTag = "[Speed: Extra Fast]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchExtraFastMixedCase()
        {
            var testTag = "[Speed: extra fast]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void DoesntMatchWithNoParameters()
        {
            var testTag = "[Speed]";
            var match = tagRegex.Match(testTag);

            Assert.IsFalse(match.Success);

            testTag = "[Speed:]";
            match = tagRegex.Match(testTag);

            Assert.IsFalse(match.Success);

            testTag = "[Speed: ]";
            match = tagRegex.Match(testTag);

            Assert.IsFalse(match.Success);
        }

        [TestMethod]
        public void MatchListSingleInSentence()
        {
            var sentence = "This is [speed: fast]a test[endspeed].";
            var matches = new StartSpeedTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 1, "More than one match.");

            var match = matches[0];
            Assert.IsTrue(match.Start == 8, "Match start was incorrect.");
            Assert.IsTrue(match.End == 20, "Match end was incorrect.");
        }

        [TestMethod]
        public void MatchListMultipleInSentence()
        {
            var sentence = "This [speed: slow]has multiple[endspeed][speed: fast] matches.[endspeed]";
            var matches = new StartSpeedTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 2, "Didn't match all.");
        }
    }
}
