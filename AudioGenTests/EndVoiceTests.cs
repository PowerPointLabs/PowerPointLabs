using System.Linq;
using System.Text.RegularExpressions;
using AudioGen.TagMatchers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AudioGenTests
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
        public void MatchEnd()
        {
            var testTag = "[EndVoice]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchEndLowercase()
        {
            var testTag = "[endvoice]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
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
        public void MatchListMultipleInSentence()
        {
            var sentence = "This [voice: male]has multiple[endvoice][voice: female] matches.[endvoice]";
            var matches = new EndVoiceTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 2, "Didn't match all.");
        }
    }
}
