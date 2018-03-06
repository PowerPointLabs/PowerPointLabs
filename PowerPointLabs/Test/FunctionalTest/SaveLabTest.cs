using System.Collections.Generic;
using System.Drawing;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

using PowerPointPresentation = PowerPointLabs.Models.PowerPointPresentation;
using Presentations = Microsoft.Office.Interop.PowerPoint.Presentations;
using Presentation = Microsoft.Office.Interop.PowerPoint.Presentation;

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

            // Select slides in the original test pptx using the slide index
            List<int> TestSlideIndexArray = InitialiseTestSlideIndexArray();
            for (int i = 0; i < NoOfTestSlides; i++)
            {
                PpOperations.SelectSlide(TestSlideIndexArray[i]);
                System.Diagnostics.Debug.WriteLine(PpOperations.GetCurrentSelection().SlideRange.Count);
            }
            
            // Get the current presentation and Save presentation
            PowerPointPresentation currentPresentation = new PowerPointPresentation();
            currentPresentation.Presentation = PpOperations.GetCurrentSelection().Application.ActivePresentation;
            //PowerPointPresentation currentPresentation = (PowerPointPresentation) PowerPointPresentation.Application.ActiveWindow.Presentation;
            //System.Diagnostics.Debug.WriteLine();
            //Presentation currPres = new Microsoft.Office.Interop.PowerPoint.Application().Presentations[1];
            //PowerPointPresentation currentPresentation = PowerPointLabs.ActionFramework.Common.Extension.FunctionalTestExtensions.GetCurrentPresentation();
            System.Diagnostics.Debug.WriteLine(currentPresentation.SelectedSlides.Count);
            PowerPointLabs.SaveLab.SaveLabMain.SaveFile(currentPresentation, true);

            // Wait for the presentation to be saved
            ThreadUtil.WaitFor(1500);

            // Open up the saved copy in the background
            Presentations newPres = new Microsoft.Office.Interop.PowerPoint.Application().Presentations;
            Presentation tempPres = newPres.Open(PowerPointLabs.SaveLab.SaveLabSettings.GetDefaultSavePresentationFileName(), Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
            PowerPointPresentation newPresentation = new PowerPointPresentation(tempPres);

            // Check each slide to ensure that it is the same
            List<int> ExpectedSlideIndexArray = InitialiseExpectedSlideIndexArray();
            for (int i = 0; i < NoOfTestSlides; i++)
            {
                AssertIsSame(currentPresentation, TestSlideIndexArray[i], newPresentation, ExpectedSlideIndexArray[i]);
            }

            // Delete the copied presentation
            System.IO.File.Delete(PowerPointLabs.SaveLab.SaveLabSettings.GetDefaultSavePresentationFileName());
            
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

        private void AssertIsSame(PowerPointPresentation originalPresentation, int originalSlideNo, PowerPointPresentation expectedPresentation, int expectedSlideNo)
        {
            Microsoft.Office.Interop.PowerPoint.Slide originalSlide = (Microsoft.Office.Interop.PowerPoint.Slide) originalPresentation.Slides[originalSlideNo - 1];
            Microsoft.Office.Interop.PowerPoint.Slide expectedSlide = (Microsoft.Office.Interop.PowerPoint.Slide) expectedPresentation.Slides[expectedSlideNo - 1];

            SlideUtil.IsSameLooking(expectedSlide, originalSlide);
            SlideUtil.IsSameAnimations(expectedSlide, originalSlide);
        }
    }
}
