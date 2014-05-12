using System.Linq;
using System.Text.RegularExpressions;
using AudioGen.TagMatchers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AudioGenTests
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
        public void MatchMale()
        {
            var testTag = "[Voice: Male]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchMaleLowercase()
        {
            var testTag = "[voice: male]";
            TagTestHelpers.MatchAndAssert(testTag, tagRegex);
        }

        [TestMethod]
        public void MatchListSingleInSentence()
        {
            var sentence = "This is [voice: female]a test[endvoice].";
            var matches = new StartVoiceTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 1, "More than one match.");

            var match = matches[0];
            Assert.IsTrue(match.Start == 8, "Match start was incorrect.");
            Assert.IsTrue(match.End == 22, "Match end was incorrect.");
        }

        [TestMethod]
        public void MatchListMultipleInSentence()
        {
            var sentence = "This [voice: male]has multiple[endvoice][voice: female] matches.[endvoice]";
            var matches = new StartVoiceTagMatcher().Matches(sentence);

            Assert.IsTrue(matches.Any(), "No matches found.");
            Assert.IsTrue(matches.Count == 2, "Didn't match all.");
        }
    }
}
