using System;
using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointLabs.ActionFramework.Common.Extension;
using Test.Util;

namespace Test.FunctionalTest
{
    [TestClass]
    public class AgendaLabBeamSyncTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AgendaLab\\AgendaSlidesBeamBeforeSync.pptx";
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AgendaLabBeamSyncTest()
        {
            BeamSyncSuccessful();
            NoHighlightedTextUnsuccessful();
            NoBeamUnsuccessful();
            NoRefSlideUnsuccessful();
        }

        [TestMethod]
        [TestCategory("FT")]
        public void FT_AgendaLabBeamSyncTest2()
        {
            NoBeamTextUnsuccessful();
            NoAgendaUnsuccessful();
        }

        public void BeamSyncSuccessful()
        {
            // TODO: Is there really no way to programmatically select multiple slides at once?
            // TODO: Ideally, I want to select the slides 5,6,7,8 together and apply Synchronise Agenda on them

            ClickOnSlideThumbnailsPanel();
            PpOperations.SelectSlide(5);
            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Reorganise Sidebar",
                "The sections have been changed. Do you wish to reorganise the items in the sidebar?",
                PplFeatures.SynchronizeAgenda,
                "&Yes");

            PpOperations.SelectSlide(6);
            PplFeatures.SynchronizeAgenda();

            ClickOnSlideThumbnailsPanel();
            PpOperations.SelectSlide(8);
            PplFeatures.SynchronizeAgenda();

            List<TestInterface.ISlideData> actualSlides = PpOperations.FetchCurrentPresentationData();
            List<TestInterface.ISlideData> expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaLab\\AgendaSlidesBeamAfterSync.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }

        public void NoHighlightedTextUnsuccessful()
        {
            PpOperations.SelectSlide(1);
            Microsoft.Office.Interop.PowerPoint.Shape highlightedText = PpOperations.RecursiveGetShapeWithPrefix("PptLabsAgenda_&^@BeamShapeMainGroup",              
                                                                           "PptLabsAgenda_&^@BeamShapeHighlightedText");
            highlightedText.SafeDelete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "The current reference slide is invalid. Either replace the reference slide or regenerate the agenda.",
                PplFeatures.SynchronizeAgenda);
        }

        public void NoBeamUnsuccessful()
        {
            PpOperations.SelectSlide(1);
            Microsoft.Office.Interop.PowerPoint.Shape beamShape = PpOperations.SelectShapesByPrefix("PptLabsAgenda_&^@BeamShapeMainGroup")[1];
            beamShape.SafeDelete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "The current reference slide is invalid. Either replace the reference slide or regenerate the agenda.",
                PplFeatures.SynchronizeAgenda);
        }

        public void NoRefSlideUnsuccessful()
        {
            Microsoft.Office.Interop.PowerPoint.Slide refSlide = PpOperations.SelectSlide(1);
            refSlide.Delete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "No reference slide could be found. Either replace the reference slide or regenerate the agenda.",
                PplFeatures.SynchronizeAgenda);
        }

        public void NoBeamTextUnsuccessful()
        {
            PpOperations.SelectSlide(1);
            for (int i = 0; i < 5; ++i)
            {
                Microsoft.Office.Interop.PowerPoint.Shape beamText = PpOperations.RecursiveGetShapeWithPrefix("PptLabsAgenda_&^@BeamShapeMainGroup",
                    "PptLabsAgenda_&^@BeamShapeText");
                beamText.SafeDelete();
            }

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "The current reference slide is invalid. Either replace the reference slide or regenerate the agenda.",
                PplFeatures.SynchronizeAgenda);
        }

        public void NoAgendaUnsuccessful()
        {
            PplFeatures.RemoveAgenda();
            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "There is no generated agenda.",
                PplFeatures.SynchronizeAgenda);
        }

        // Click Thumbnails Panel to make selectedSlide focused.
        // When focused, add the agenda beam to the selectedSlide (doesn't have agenda) & sync.
        // When unfocused, only sync (so selectedSlide (doesn't have agenda) remains the same).
        private static void ClickOnSlideThumbnailsPanel()
        {
            IntPtr pptPanel = NativeUtil.FindWindow("PPTFrameClass", null);
            IntPtr mdiPanel = NativeUtil.FindWindowEx(pptPanel, IntPtr.Zero, "MDIClient", null);
            IntPtr mdiPanel2 = NativeUtil.FindWindowEx(mdiPanel, IntPtr.Zero, "mdiClass", null);
            if (PpOperations.IsOffice2010())
            {
                IntPtr thumbnailsPanel = NativeUtil.FindWindowEx(mdiPanel2, IntPtr.Zero, "paneClassDC", "Thumbnails");
                NativeUtil.SendMessage(thumbnailsPanel, 0x0201 /*left button down*/, IntPtr.Zero, IntPtr.Zero);
            } 
            else // Office2013 or Higher
            {
                NativeUtil.SendMessage(mdiPanel2, 0x0201 /*left button down*/, IntPtr.Zero, IntPtr.Zero);
            }
        }
    }
}