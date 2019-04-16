using System.Linq;
using System.Text.RegularExpressions;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.TagMatchers;

using Test.Util;

namespace Test.UnitTest.TagMatchers
{
    [TestClass]
    public class PronounceTests
    {
        private Regex tagRegex;

        [TestInitialize]
        public void Initialize()
        {
            tagRegex = new PronounceTagMatcher().Regex;
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchPronounceBlock()
        {
            string testTag = "[Pronounce: <IPA>]<Word Here>[EndPronounce]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchPronounceBlockCaseInsensitive()
        {
            string testTag = "[pronounce: <IPA>]<word here>[endpronounce]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchListSingleInSentence()
        {
            string sentence = "This is a [pronounce: <IPA>]test[endpronounce].";
            System.Collections.Generic.List<PowerPointLabs.Tags.ITag> matches = new PronounceTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 1, "More than one match.");

            PowerPointLabs.Tags.ITag match = matches[0];
            Assert.IsTrue(match.Start == 10, "Match start was incorrect.");
            Assert.IsTrue(match.End == 45, "Match end was incorrect.");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchListMultipleInSentence()
        {
            string sentence = "This [pronounce: <IPA>]has[endpronounce] multiple[pause: 2][pronounce: <IPA>]matches.[endpronounce]";
            System.Collections.Generic.List<PowerPointLabs.Tags.ITag> matches = new PronounceTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 2, "Didn't match all.");
        }
    }
}
