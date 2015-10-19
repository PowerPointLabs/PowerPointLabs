﻿using System;
using System.Linq;
using FunctionalTest.util;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace FunctionalTest
{
    [TestClass]
    public class AgendaLabBeamSyncTest : BaseFunctionalTest
    {
        protected override string GetTestingSlideName()
        {
            return "AgendaSlidesBeamBeforeSync.pptx";
        }

        [TestMethod]
        public void FT_AgendaLabBeamSyncTest()
        {
            BeamSyncSuccessful();
            NoHighlightedTextUnsuccessful();
            NoBeamUnsuccessful();
            NoRefSlideUnsuccessful();
        }

        [TestMethod]
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
            PplFeatures.SynchronizeAgenda();

            PpOperations.SelectSlide(6);
            PplFeatures.SynchronizeAgenda();

            ClickOnSlideThumbnailsPanel();
            PpOperations.SelectSlide(8);
            PplFeatures.SynchronizeAgenda();

            var actualSlides = PpOperations.FetchCurrentPresentationData();
            var expectedSlides = PpOperations.FetchPresentationData(
                PathUtil.GetDocTestPresentationPath("AgendaSlidesBeamAfterSync.pptx"));
            PresentationUtil.AssertEqual(expectedSlides, actualSlides);
        }

        public void NoHighlightedTextUnsuccessful()
        {
            PpOperations.SelectSlide(1);
            var highlightedText = PpOperations.RecursiveGetShapeWithPrefix("PptLabsAgenda_&^@BeamShapeMainGroup",              
                                                                           "PptLabsAgenda_&^@BeamShapeHighlightedText");
            highlightedText.Delete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "The current reference slide is invalid. Either replace the reference slide or regenerate the agenda.",
                PplFeatures.SynchronizeAgenda);
        }

        public void NoBeamUnsuccessful()
        {
            PpOperations.SelectSlide(1);
            var beamShape = PpOperations.SelectShapesByPrefix("PptLabsAgenda_&^@BeamShapeMainGroup")[1];
            beamShape.Delete();

            MessageBoxUtil.ExpectMessageBoxWillPopUp(
                "Unable to execute action",
                "The current reference slide is invalid. Either replace the reference slide or regenerate the agenda.",
                PplFeatures.SynchronizeAgenda);
        }

        public void NoRefSlideUnsuccessful()
        {
            var refSlide = PpOperations.SelectSlide(1);
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
                var beamText = PpOperations.RecursiveGetShapeWithPrefix("PptLabsAgenda_&^@BeamShapeMainGroup",
                    "PptLabsAgenda_&^@BeamShapeText");
                beamText.Delete();
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
            var pptPanel = NativeUtil.FindWindow("PPTFrameClass", null);
            var mdiPanel = NativeUtil.FindWindowEx(pptPanel, IntPtr.Zero, "MDIClient", null);
            var mdiPanel2 = NativeUtil.FindWindowEx(mdiPanel, IntPtr.Zero, "mdiClass", null);
            if (PpOperations.IsOffice2010())
            {
                var thumbnailsPanel = NativeUtil.FindWindowEx(mdiPanel2, IntPtr.Zero, "paneClassDC", "Thumbnails");
                NativeUtil.SendMessage(thumbnailsPanel, 0x0201 /*left button down*/, IntPtr.Zero, IntPtr.Zero);
            } 
            else if (PpOperations.IsOffice2013())
            {
                NativeUtil.SendMessage(mdiPanel2, 0x0201 /*left button down*/, IntPtr.Zero, IntPtr.Zero);
            }
        }
    }
}