using System.Linq;
using System.Text.RegularExpressions;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.TagMatchers;

using Test.Util;

namespace Test.UnitTest.TagMatchers
{
    [TestClass]
    public class StartVoiceTests
    {
        private Regex tagRegex;

        [TestInitialize]
        public void Initialize()
        {
            tagRegex = new StartVoiceTagMatcher().Regex;
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchMale()
        {
            string testTag = "[Voice: Male]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchMaleLowercase()
        {
            string testTag = "[voice: male]";
            TagUtil.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchListSingleInSentence()
        {
            string sentence = "This is [voice: female]a test[endvoice].";
            System.Collections.Generic.List<PowerPointLabs.Tags.ITag> matches = new StartVoiceTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 1, "More than one match.");

            PowerPointLabs.Tags.ITag match = matches[0];
            Assert.IsTrue(match.Start == 8, "Match start was incorrect.");
            Assert.IsTrue(match.End == 22, "Match end was incorrect.");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void MatchListMultipleInSentence()
        {
            string sentence = "This [voice: male]has multiple[endvoice][voice: female] matches.[endvoice]";
            System.Collections.Generic.List<PowerPointLabs.Tags.ITag> matches = new StartVoiceTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 2, "Didn't match all.");
        }
    }
}
