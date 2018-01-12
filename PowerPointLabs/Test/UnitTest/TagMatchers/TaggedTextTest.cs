using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.Models;

namespace Test.UnitTest.TagMatchers
{
    [TestClass]
    public class TaggedTextTest
    {
        private const string SentenceWithTags = "This has [speed: slow]some tags[endspeed] [voice: mike]in it.[endvoice]";

        [TestMethod]
        [TestCategory("UT")]
        public void SplitStringsByClick()
        {
            const string sentence = "This is separated by a click.[afterclick]This is the next part.";

            var t = new TaggedText(sentence);

            var split = t.SplitByClicks();
            Assert.IsTrue(split.Count == 2, "Split into incorrect amount of strings.");
            Assert.IsTrue(split[0].Equals("This is separated by a click."), "First split incorrect.");
            Assert.IsTrue(split[1].Equals("This is the next part."), "Second split incorrect.");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void LeaveStringIntactWhenNoSplit()
        {
            const string sentence = "This has no clicks to split by.";

            var t = new TaggedText(sentence);
            var split = t.SplitByClicks();

            Assert.IsTrue(split.Count == 1, "Split when there wasn't a click.");
            Assert.IsTrue(split[0].Equals("This has no clicks to split by."), "Didn't leave original string intact.");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void RemoveTagsFromText()
        {
            const string expected = "This has some tags in it.";

            var t = new TaggedText(SentenceWithTags);
            var actual = t.ToPrettyString();

            Assert.AreEqual(expected, actual, "Didn't remove tags properly.");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TestToString()
        {
            var expected = SentenceWithTags;
            
            var t = new TaggedText(SentenceWithTags);
            var actual = t.ToString();

            Assert.AreEqual(expected, actual, "Didn't produce the original sentence.");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void SplitEmptyString()
        {
            var t = new TaggedText("");
            var result = t.SplitByClicks();

            Assert.IsNotNull(result, "Returned a null list.");
            Assert.IsFalse(result.Any(), "List contained results.");
        }
    }
}
