using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.TagMatchers;
using Test.Util;

namespace Test.UnitTest
{
    [TestClass]
    public class EndVoiceTests
    {
        private Regex tagRegex;

        [TestInitialize]
        public void Initialize()
        {
            tagRegex = new EndVoiceTagMatcher().Regex;
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchEnd()
        {
            var testTag = "[EndVoice]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchEndLowercase()
        {
            var testTag = "[endvoice]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchListSingleInSentence()
        {
            var sentence = "This is [voice: female]a test[endvoice].";
            var matches = new EndVoiceTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 1, "More than one match.");

            var match = matches[0];
            Assert.IsTrue(match.Start == 29, "Match start was incorrect.");
            Assert.IsTrue(match.End == 38, "Match end was incorrect.");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchListMultipleInSentence()
        {
            var sentence = "This [voice: male]has multiple[endvoice][voice: female] matches.[endvoice]";
            var matches = new EndVoiceTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 2, "Didn't match all.");
        }
    }
}
