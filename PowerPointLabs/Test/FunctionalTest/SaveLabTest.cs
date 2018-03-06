using System.Collections.Generic;
using System.Drawing;

using Microsoft.Office.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using Test.Util;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

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

        private const string SavedSlideName = "SaveLab\\SaveLab_Copy.pptx";

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
            }

            // Get the current presentation and Save presentation
            PowerPointLabs.Models.PowerPointPresentation currentPresentation = (PowerPointLabs.Models.PowerPointPresentation) PowerPointLabs.Models.PowerPointPresentation.Application.ActivePresentation;
            PowerPointLabs.SaveLab.SaveLabMain.SaveFile(currentPresentation, true);

            // Wait for the presentation to be saved

            // Open up the saved copy in the background

            // Check each slide to ensure that it is the same

            // Delete the copied presentation
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

        private void RightClick(Shape target)
        {
            Point pt = new Point(
                PpOperations.PointsToScreenPixelsX(target.Left + target.Width / 2),
                PpOperations.PointsToScreenPixelsY(target.Top + target.Height / 2));
            MouseUtil.SendMouseRightClick(pt.X, pt.Y);
        }
    }
}
