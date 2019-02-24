﻿using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.Office.Tools;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.NarrationsLab;
using PowerPointLabs.NarrationsLab.Data;
using PowerPointLabs.NarrationsLab.Views;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.NarrationsLab
{
    [ExportActionRibbonId(NarrationsLabText.AddNarrationsTag)]
    class AddNarrationsActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            if (NotesToAudio.IsAzureVoiceSelected && UserAccount.GetInstance().IsEmpty())
            {
                MessageBox.Show("Invalid user account. Please log in again.");
                return;
            }

            //TODO: This needs to improved to stop using global variables
            this.StartNewUndoEntry();

            PowerPointSlide currentSlide = this.GetCurrentSlide();

            // If there are text in notes page for any of the selected slides 
            if (this.GetCurrentPresentation().SelectedSlides.Any(slide => slide.NotesPageText.Trim() != ""))
            {
                NotesToAudio.IsRemoveAudioEnabled = true;
                this.GetRibbonUi().RefreshRibbonControl("RemoveNarrationsButton");
            }
            List<string[]> allAudioFiles = new List<string[]>();
            try
            {
                allAudioFiles = NotesToAudio.EmbedSelectedSlideNotes();
            }
            catch
            {
                MessageBox.Show("Failed to generate audio files.");
                return;
            }

            CustomTaskPane recorderPane = this.GetAddIn().GetActivePane(typeof(RecorderTaskPane));

            if (recorderPane == null)
            {
                return;
            }

            RecorderTaskPane recorder = recorderPane.Control as RecorderTaskPane;

            if (recorder == null)
            {
                return;
            }

            // initialize selected slides' audio
            recorder.InitializeAudioAndScript(this.GetCurrentPresentation().SelectedSlides.ToList(),
                                                  allAudioFiles, true);

            // if current list is visible, update the pane immediately
            if (recorderPane.Visible)
            {
                recorder.UpdateLists(currentSlide.ID);
            }

            if (NarrationsLabSettings.IsPreviewEnabled)
            {
                NotesToAudio.PreviewAnimations();
            }
        }
    }
}
