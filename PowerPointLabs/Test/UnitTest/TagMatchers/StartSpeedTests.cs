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
            string testTag = "[speed: extra slow]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchExtraSlowUpperCase()
        {
            string testTag = "[Speed: Extra Slow]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchExtraSlowMixedCase()
        {
            string testTag = "[Speed: extra Slow]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchSlowLowerCase()
        {
            string testTag = "[speed: slow]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchSlowUpperCase()
        {
            string testTag = "[Speed: Slow]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchSlowMixedCase()
        {
            string testTag = "[Speed: slow]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchMediumLowerCase()
        {
            string testTag = "[speed: medium]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchMediumUpperCase()
        {
            string testTag = "[Speed: Medium]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchMediumMixedCase()
        {
            string testTag = "[Speed: medium]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchFastLowerCase()
        {
            string testTag = "[speed: fast]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchFastUpperCase()
        {
            string testTag = "[Speed: Fast]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchFastMixedCase()
        {
            string testTag = "[Speed: fast]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchExtraFastLowerCase()
        {
            string testTag = "[speed: extra fast]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchExtraFastUpperCase()
        {
            string testTag = "[Speed: Extra Fast]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchExtraFastMixedCase()
        {
            string testTag = "[Speed: extra fast]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void DoesntMatchWithNoParameters()
        {
            string testTag = "[Speed]";
            Match match = tagRegex.Match(testTag);

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
            string sentence = "This is [speed: fast]a test[endspeed].";
            System.Collections.Generic.List<PowerPointLabs.Tags.ITag> matches = new StartSpeedTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 1, "More than one match.");

            PowerPointLabs.Tags.ITag match = matches[0];
            Assert.IsTrue(match.Start == 8, "Match start was incorrect.");
            Assert.IsTrue(match.End == 20, "Match end was incorrect.");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchListMultipleInSentence()
        {
            string sentence = "This [speed: slow]has multiple[endspeed][speed: fast] matches.[endspeed]";
            System.Collections.Generic.List<PowerPointLabs.Tags.ITag> matches = new StartSpeedTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 2, "Didn't match all.");
        }
    }
}
