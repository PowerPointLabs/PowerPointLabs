using System;
using System.Collections.Generic;
using System.Linq;
using FunctionalTest.models;
using FunctionalTest.util;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SlideData = FunctionalTest.models.PresentationCompareData.SlideData;

namespace FunctionalTest
{
    [TestClass]
    public class AgendaLabGenerateTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AgendaSlidesVisualDefault.pptx";
        }

        [TestMethod]
        public void FT_AgendaLabGenerate()
        {
            // AgendaSlidesVisualDefault -> Remove ->
            // AgendaSlidesDefault -> Generate Text ->
            // AgendaSlidesTextDefault -> Generate Beam ->
            // AgendaSlidesBeamDefault -> Generate Visual ->
            // AgendaSlidesVisualDefault
            var visualDefaultSlides = SaveAllSlides();

            PplFeatures.RemoveAgenda();
            var actualSlides = SaveAllSlides();
            OpenAnotherPresentation("AgendaSlidesDefault.pptx");
            var expectedSlides = SaveAllSlides();
            AssertEqualSlides(expectedSlides, actualSlides);

            PplFeatures.GenerateTextAgenda();
            actualSlides = SaveAllSlides();
            OpenAnotherPresentation("AgendaSlidesTextDefault.pptx");
            expectedSlides = SaveAllSlides();
            AssertEqualSlides(expectedSlides, actualSlides);

            //PplFeatures.GenerateBeamAgenda();
            MessageBoxUtil.ExpectMessageBoxWillPopUp("Confirm Update",
                "Agenda already exists. By confirm this dialog agenda will be regenerated. Do you want to proceed?",
                PplFeatures.GenerateBeamAgenda, buttonNameToClick: "OK");
            actualSlides = SaveAllSlides();
            OpenAnotherPresentation("AgendaSlidesBeamDefault.pptx");
            expectedSlides = SaveAllSlides();
            AssertEqualSlides(expectedSlides, actualSlides);

            //PplFeatures.GenerateVisualAgenda();
            MessageBoxUtil.ExpectMessageBoxWillPopUp("Confirm Update",
                "Agenda already exists. By confirm this dialog agenda will be regenerated. Do you want to proceed?",
                PplFeatures.GenerateVisualAgenda, buttonNameToClick: "OK");
            actualSlides = SaveAllSlides();
            AssertEqualSlides(visualDefaultSlides, actualSlides);
        }

        public List<SlideData> SaveAllSlides()
        {
            return PpOperations.GetAllSlides().Select(SlideData.SaveSlideData)
                                              .ToList();
        }

        private void AssertEqualSlides(List<SlideData> expectedSlides, List<SlideData> actualSlides)
        {   
            Assert.AreEqual(expectedSlides.Count, actualSlides.Count);
            int count = expectedSlides.Count;
            for (int i = 0; i < count; ++i)
            {
                SlideData.AssertEqual(expectedSlides[i], actualSlides[i]);
            }
        }
    }
}
