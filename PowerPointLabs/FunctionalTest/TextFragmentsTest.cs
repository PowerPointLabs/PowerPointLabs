﻿using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class TextFragmentsTest : BaseFunctionalTest
    {
        private const string ShapeName = "Text Fragments Shape";

        protected override string GetTestingSlideName()
        {
            return "TextFragments.pptx";
        }

        [TestMethod]
        public void FT_TextFragmentsTest()
        {
            PpOperations.SelectSlide(3);
            var fragments = new[]
            {
                new[] {17, 37},
                new[] {45, 114},
                new[] {122, 140},
                new[] {148, 166},
                new[] {174, 192},
                new[] {200, 218},
                new[] {226, 234},
                new[] {242, 257},
                new[] {268, 286},
                new[] {305, 313},
                new[] {370, 388},
            };

            foreach (var fragment in fragments)
            {
                PpOperations.SelectTextInShape(ShapeName, fragment[0], fragment[1]);
                PplFeatures.HighlightFragments();
            }

            AssertIsSame(3, 4);
        }

        private void AssertIsSame(int actualSlideIndex, int expectedSlideIndex)
        {
            var actualSlide = PpOperations.SelectSlide(actualSlideIndex);
            var expectedSlide = PpOperations.SelectSlide(expectedSlideIndex);
            SlideUtil.IsSameLooking(expectedSlide, actualSlide);
            SlideUtil.IsSameAnimations(expectedSlide, actualSlide);
        }

        /* // A useful utility function to find out the indexes of certain characters in the text.
        private void AnalyseTextIndexes()
        {
            PpOperations.SelectSlide(3);
            var t = PpOperations.SelectAllTextInShape(ShapeName);
            for (int i = 0; i < t.Length; ++i)
            {
                index = i + 1;
                Debug.WriteLine(index + " : " + t[i]);
            }
            Debug.WriteLine(t);
        }
        */
    }
}
