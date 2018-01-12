using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.TagMatchers;
using Test.Util;

namespace Test.UnitTest.TagMatchers
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
        [TestCategory("UT")]
        public void MatchExtraSlowLowerCase()
        {
            var testTag = "[speed: extra slow]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchExtraSlowUpperCase()
        {
            var testTag = "[Speed: Extra Slow]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchExtraSlowMixedCase()
        {
            var testTag = "[Speed: extra Slow]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchSlowLowerCase()
        {
            var testTag = "[speed: slow]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchSlowUpperCase()
        {
            var testTag = "[Speed: Slow]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchSlowMixedCase()
        {
            var testTag = "[Speed: slow]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchMediumLowerCase()
        {
            var testTag = "[speed: medium]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchMediumUpperCase()
        {
            var testTag = "[Speed: Medium]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchMediumMixedCase()
        {
            var testTag = "[Speed: medium]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchFastLowerCase()
        {
            var testTag = "[speed: fast]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchFastUpperCase()
        {
            var testTag = "[Speed: Fast]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchFastMixedCase()
        {
            var testTag = "[Speed: fast]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchExtraFastLowerCase()
        {
            var testTag = "[speed: extra fast]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchExtraFastUpperCase()
        {
            var testTag = "[Speed: Extra Fast]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchExtraFastMixedCase()
        {
            var testTag = "[Speed: extra fast]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
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
        [TestCategory("UT")]
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
        [TestCategory("UT")]
        public void MatchListMultipleInSentence()
        {
            var sentence = "This [speed: slow]has multiple[endspeed][speed: fast] matches.[endspeed]";
            var matches = new StartSpeedTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 2, "Didn't match all.");
        }
    }
}
