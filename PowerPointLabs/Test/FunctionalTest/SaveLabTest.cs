using System.Collections.Generic;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

using PowerPointPresentation = PowerPointLabs.Models.PowerPointPresentation;

namespace Test.FunctionalTest
{
    [TestClass]
    public class SaveLabTest : BaseFunctionalTest
    {
        //Number of test slides
        private const int NoOfTestSlides = 4;

        //Slide Numbers
        private const int OriginalTextboxSlideSlideNo = 1;
        private const int ExpectedTextboxSlideSlideNo = 1;
        private const int OriginalShapesSlideSlideNo = 3;
        private const int ExpectedShapesSlideSlideNo = 2;
        private const int OriginalPicturesSlideSlideNo = 5;
        private const int ExpectedPicturesSlideSlideNo = 3;
        private const int OriginalAnimationSlideSlideNo = 7;
        private const int ExpectedAnimationSlideSlideNo = 4;

        protected override string GetTestingSlideName()
        {
            return "SaveLab\\SaveLab.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_SaveLabTest()
        {
            // Get the full path of the expected slides
            string ExpectedSlidesFullName = System.IO.Path.Combine(PathUtil.GetDocTestPath(), "SaveLab\\SaveLab_Copy.pptx");
            
            if (System.IO.File.Exists(ExpectedSlidesFullName))
            {
                // Gather slide indexes for both original and expected slides
                List<int> TestSlideIndexArray = InitialiseTestSlideIndexArray();
                List<int> ExpectedSlideIndexArray = InitialiseExpectedSlideIndexArray();

                // Open up the saved copy in the background
                Presentations newPres = new Microsoft.Office.Interop.PowerPoint.Application().Presentations;
                Presentation tempPres = newPres.Open(ExpectedSlidesFullName, WithWindow: MsoTriState.msoFalse);
                PowerPointPresentation newPresentation = new PowerPointPresentation(tempPres);
                
                // Check each slide to ensure that it is the same
                for (int i = 0; i < NoOfTestSlides; i++)
                {
                    AssertIsSame(TestSlideIndexArray[i], ExpectedSlideIndexArray[i], tempPres);
                }

                // Close the saved copy
                newPresentation.Close();
                
            }
                
        }

        private List<int> InitialiseTestSlideIndexArray()
        {
            List<int> indexArray = new List<int>();
            indexArray.Add(OriginalTextboxSlideSlideNo);
            indexArray.Add(OriginalShapesSlideSlideNo);
            indexArray.Add(OriginalPicturesSlideSlideNo);
            indexArray.Add(OriginalAnimationSlideSlideNo);

            return indexArray;
        }

        private List<int> InitialiseExpectedSlideIndexArray()
        {
            List<int> indexArray = new List<int>();
            indexArray.Add(ExpectedTextboxSlideSlideNo);
            indexArray.Add(ExpectedShapesSlideSlideNo);
            indexArray.Add(ExpectedPicturesSlideSlideNo);
            indexArray.Add(ExpectedAnimationSlideSlideNo);

            return indexArray;
        }

        private void AssertIsSame(int originalSlideNo, int expectedSlideNo, Presentation expectedPresentation)
        {
            Microsoft.Office.Interop.PowerPoint.Slide originalSlide = PpOperations.SelectSlide(originalSlideNo);
            Microsoft.Office.Interop.PowerPoint.Slide expectedSlide = expectedPresentation.Slides[expectedSlideNo];

            SlideUtil.IsSameLooking(expectedSlide, originalSlide);
            SlideUtil.IsSameAnimations(expectedSlide, originalSlide);
        }
    }
}
